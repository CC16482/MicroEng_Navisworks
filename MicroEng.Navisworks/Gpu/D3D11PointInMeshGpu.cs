using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using SharpDX;
using SharpDX.Direct3D;
using SharpDX.Direct3D11;
using SharpDX.D3DCompiler;
using Buffer = SharpDX.Direct3D11.Buffer;

namespace MicroEng.Navisworks.SpaceMapper.Gpu
{
    internal sealed class D3D11PointInMeshGpu : IDisposable, IPointInMeshGpuBackend
    {
        // Result encoding:
        // 0 = outside
        // 1 = inside
        // 2 = uncertain (ray degeneracy / mismatch in dual-ray mode)
        public const uint Outside = 0;
        public const uint Inside = 1;
        public const uint Uncertain = 2;

        internal sealed class AdapterInfo
        {
            public string Description { get; set; }
            public int VendorId { get; set; }
            public int DeviceId { get; set; }
            public int SubsystemId { get; set; }
            public int Revision { get; set; }
            public long DedicatedVideoMemory { get; set; }
            public long DedicatedSystemMemory { get; set; }
            public long SharedSystemMemory { get; set; }
            public string Luid { get; set; }
            public string FeatureLevel { get; set; }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct Constants
        {
            public uint TriangleCount;
            public uint PointCount;
            public uint ZoneCount;
            public uint UseSecondRay; // 0/1
        }

        private readonly Device _device;
        private readonly DeviceContext _ctx;
        private readonly ComputeShader _cs;

        private Buffer _cbConstants;

        private Buffer _triBuffer;
        private ShaderResourceView _triSrv;
        private int _triCapacity;

        private Buffer _ptBuffer;
        private ShaderResourceView _ptSrv;
        private int _ptCapacity;

        private Buffer _pointZoneBuffer;
        private ShaderResourceView _pointZoneSrv;
        private int _pointZoneCapacity;

        private Buffer _zoneRangesBuffer;
        private ShaderResourceView _zoneRangesSrv;
        private int _zoneRangesCapacity;

        private Buffer _outBuffer;
        private UnorderedAccessView _outUav;
        private Buffer _outStaging;
        private int _outCapacity;

        private const int Threads = 256;

        public AdapterInfo GpuAdapter { get; }
        public string BackendName => "D3D11";
        public string DeviceName => GpuAdapter?.Description ?? "D3D11 device";
        public TimeSpan LastDispatchTime { get; private set; }
        public TimeSpan LastReadbackTime { get; private set; }

        private D3D11PointInMeshGpu(Device device, ComputeShader cs)
        {
            _device = device;
            _ctx = device.ImmediateContext;
            _cs = cs;
            GpuAdapter = TryGetAdapterInfo(device);

            _cbConstants = new Buffer(_device,
                Utilities.SizeOf<Constants>(),
                ResourceUsage.Dynamic,
                BindFlags.ConstantBuffer,
                CpuAccessFlags.Write,
                ResourceOptionFlags.None,
                0);
        }

        public static bool TryCreate(out D3D11PointInMeshGpu gpu, out string reason)
        {
            gpu = null;
            reason = null;

            try
            {
                var device = new Device(DriverType.Hardware, DeviceCreationFlags.None, FeatureLevel.Level_11_0);
                if (device.FeatureLevel < FeatureLevel.Level_11_0)
                {
                    reason = $"D3D11 feature level {device.FeatureLevel} does not support compute shaders.";
                    device.Dispose();
                    return false;
                }

                using (var bytecode = ShaderBytecode.Compile(ShaderSource, "CSMain", "cs_5_0",
                    ShaderFlags.OptimizationLevel3, EffectFlags.None))
                {
                    var cs = new ComputeShader(device, bytecode);
                    gpu = new D3D11PointInMeshGpu(device, cs);
                }

                return true;
            }
            catch (Exception ex)
            {
                reason = ex.Message;
                return false;
            }
        }

        private static AdapterInfo TryGetAdapterInfo(Device device)
        {
            if (device == null)
            {
                return null;
            }

            try
            {
                using (var dxgiDevice = device.QueryInterface<SharpDX.DXGI.Device>())
                using (var adapter = dxgiDevice.Adapter)
                {
                    var desc = adapter.Description;
                    return new AdapterInfo
                    {
                        Description = desc.Description,
                        VendorId = desc.VendorId,
                        DeviceId = desc.DeviceId,
                        SubsystemId = desc.SubsystemId,
                        Revision = desc.Revision,
                        DedicatedVideoMemory = desc.DedicatedVideoMemory,
                        DedicatedSystemMemory = desc.DedicatedSystemMemory,
                        SharedSystemMemory = desc.SharedSystemMemory,
                        Luid = FormatLuid(desc.Luid),
                        FeatureLevel = device.FeatureLevel.ToString()
                    };
                }
            }
            catch
            {
                return null;
            }
        }

        private static string FormatLuid(object luid)
        {
            if (luid == null)
            {
                return null;
            }

            try
            {
                var type = luid.GetType();
                var highProp = type.GetProperty("HighPart");
                var lowProp = type.GetProperty("LowPart");
                if (highProp != null && lowProp != null)
                {
                    var high = Convert.ToInt32(highProp.GetValue(luid, null));
                    var low = Convert.ToUInt32(lowProp.GetValue(luid, null));
                    return $"0x{high:x8}:0x{low:x8}";
                }

                if (luid is long value)
                {
                    var high = (int)(value >> 32);
                    var low = (uint)(value & 0xffffffff);
                    return $"0x{high:x8}:0x{low:x8}";
                }
            }
            catch
            {
                // ignore format failures
            }

            return luid.ToString();
        }

        public uint[] TestPoints(
            Triangle[] trianglesLocal,
            Float4[] pointsLocal,
            bool intensiveTwoRays,
            CancellationToken token)
        {
            var result = TestPoints(
                trianglesLocal,
                pointsLocal,
                intensiveTwoRays,
                token,
                out var dispatchTime,
                out var readbackTime);

            LastDispatchTime = dispatchTime;
            LastReadbackTime = readbackTime;
            return result;
        }

        public uint[] TestPoints(
            Triangle[] trianglesLocal,
            Float4[] pointsLocal,
            bool intensiveTwoRays,
            CancellationToken token,
            out TimeSpan dispatchTime,
            out TimeSpan readbackTime)
        {
            if (trianglesLocal == null) throw new ArgumentNullException(nameof(trianglesLocal));
            if (pointsLocal == null) throw new ArgumentNullException(nameof(pointsLocal));

            if (pointsLocal.Length == 0)
            {
                dispatchTime = TimeSpan.Zero;
                readbackTime = TimeSpan.Zero;
                return Array.Empty<uint>();
            }

            var zoneRanges = new ZoneRange[1];
            zoneRanges[0] = new ZoneRange
            {
                TriStart = 0,
                TriCount = (uint)trianglesLocal.Length
            };

            var pointZoneIds = new uint[pointsLocal.Length];
            return TestPointsBatched(
                trianglesLocal,
                pointsLocal,
                pointZoneIds,
                zoneRanges,
                intensiveTwoRays,
                token,
                out dispatchTime,
                out readbackTime);
        }

        public uint[] TestPointsBatched(
            Triangle[] trianglesAll,
            Float4[] pointsAll,
            uint[] pointZoneIds,
            ZoneRange[] zoneRanges,
            bool intensiveTwoRays,
            CancellationToken token,
            out TimeSpan dispatchTime,
            out TimeSpan readbackTime)
        {
            if (trianglesAll == null) throw new ArgumentNullException(nameof(trianglesAll));
            if (pointsAll == null) throw new ArgumentNullException(nameof(pointsAll));
            if (pointZoneIds == null) throw new ArgumentNullException(nameof(pointZoneIds));
            if (zoneRanges == null) throw new ArgumentNullException(nameof(zoneRanges));
            if (pointsAll.Length != pointZoneIds.Length) throw new ArgumentException("pointZoneIds must match pointsAll length.", nameof(pointZoneIds));

            token.ThrowIfCancellationRequested();

            if (pointsAll.Length == 0)
            {
                dispatchTime = TimeSpan.Zero;
                readbackTime = TimeSpan.Zero;
                return Array.Empty<uint>();
            }

            EnsureTriangleCapacity(trianglesAll.Length);
            EnsurePointCapacity(pointsAll.Length);
            EnsurePointZoneCapacity(pointsAll.Length);
            EnsureZoneRangeCapacity(zoneRanges.Length);
            EnsureOutputCapacity(pointsAll.Length);

            UpdateDynamicStructured(_triBuffer, trianglesAll);
            UpdateDynamicStructured(_ptBuffer, pointsAll);
            UpdateDynamicStructured(_pointZoneBuffer, pointZoneIds);
            UpdateDynamicStructured(_zoneRangesBuffer, zoneRanges);

            var c = new Constants
            {
                TriangleCount = (uint)trianglesAll.Length,
                PointCount = (uint)pointsAll.Length,
                ZoneCount = (uint)zoneRanges.Length,
                UseSecondRay = intensiveTwoRays ? 1u : 0u
            };
            UpdateConstants(c);

            _ctx.ComputeShader.Set(_cs);
            _ctx.ComputeShader.SetConstantBuffer(0, _cbConstants);
            _ctx.ComputeShader.SetShaderResource(0, _triSrv);
            _ctx.ComputeShader.SetShaderResource(1, _ptSrv);
            _ctx.ComputeShader.SetShaderResource(2, _pointZoneSrv);
            _ctx.ComputeShader.SetShaderResource(3, _zoneRangesSrv);
            _ctx.ComputeShader.SetUnorderedAccessView(0, _outUav);

            var dispatchStart = Stopwatch.GetTimestamp();
            int groups = (pointsAll.Length + Threads - 1) / Threads;
            _ctx.Dispatch(groups, 1, 1);
            dispatchTime = ToTimeSpan(Stopwatch.GetTimestamp() - dispatchStart);

            _ctx.ComputeShader.SetUnorderedAccessView(0, null);
            _ctx.ComputeShader.SetShaderResource(0, null);
            _ctx.ComputeShader.SetShaderResource(1, null);
            _ctx.ComputeShader.SetShaderResource(2, null);
            _ctx.ComputeShader.SetShaderResource(3, null);
            _ctx.ComputeShader.Set(null);

            var readbackStart = Stopwatch.GetTimestamp();
            _ctx.CopyResource(_outBuffer, _outStaging);

            var result = new uint[pointsAll.Length];
            DataStream ds;
            _ctx.MapSubresource(_outStaging, 0, MapMode.Read, MapFlags.None, out ds);
            ds.ReadRange(result, 0, result.Length);
            _ctx.UnmapSubresource(_outStaging, 0);
            ds.Dispose();
            readbackTime = ToTimeSpan(Stopwatch.GetTimestamp() - readbackStart);

            return result;
        }

        private void EnsureTriangleCapacity(int triCount)
        {
            if (triCount <= _triCapacity) return;

            DisposeTriangles();

            _triCapacity = NextPow2(triCount);

            _triBuffer = CreateDynamicStructuredBuffer<Triangle>(_triCapacity, BindFlags.ShaderResource);
            _triSrv = new ShaderResourceView(_device, _triBuffer);
        }

        private void EnsurePointCapacity(int pointCount)
        {
            if (pointCount <= _ptCapacity) return;

            DisposePoints();

            _ptCapacity = NextPow2(pointCount);

            _ptBuffer = CreateDynamicStructuredBuffer<Float4>(_ptCapacity, BindFlags.ShaderResource);
            _ptSrv = new ShaderResourceView(_device, _ptBuffer);
        }

        private void EnsurePointZoneCapacity(int pointCount)
        {
            if (pointCount <= _pointZoneCapacity) return;

            DisposePointZones();

            _pointZoneCapacity = NextPow2(pointCount);

            _pointZoneBuffer = CreateDynamicStructuredBuffer<uint>(_pointZoneCapacity, BindFlags.ShaderResource);
            _pointZoneSrv = new ShaderResourceView(_device, _pointZoneBuffer);
        }

        private void EnsureZoneRangeCapacity(int zoneCount)
        {
            if (zoneCount <= _zoneRangesCapacity) return;

            DisposeZoneRanges();

            _zoneRangesCapacity = NextPow2(zoneCount);

            _zoneRangesBuffer = CreateDynamicStructuredBuffer<ZoneRange>(_zoneRangesCapacity, BindFlags.ShaderResource);
            _zoneRangesSrv = new ShaderResourceView(_device, _zoneRangesBuffer);
        }

        private void EnsureOutputCapacity(int pointCount)
        {
            if (pointCount <= _outCapacity) return;

            DisposeOutput();

            _outCapacity = NextPow2(pointCount);

            _outBuffer = CreateDefaultStructuredBuffer<uint>(_outCapacity, BindFlags.UnorderedAccess);
            _outUav = new UnorderedAccessView(_device, _outBuffer);

            _outStaging = new Buffer(_device, new BufferDescription
            {
                SizeInBytes = sizeof(uint) * _outCapacity,
                Usage = ResourceUsage.Staging,
                BindFlags = BindFlags.None,
                CpuAccessFlags = CpuAccessFlags.Read,
                OptionFlags = ResourceOptionFlags.None,
                StructureByteStride = 0
            });
        }

        private Buffer CreateDynamicStructuredBuffer<T>(int elementCount, BindFlags bind) where T : struct
        {
            return new Buffer(_device, new BufferDescription
            {
                SizeInBytes = Utilities.SizeOf<T>() * elementCount,
                Usage = ResourceUsage.Dynamic,
                BindFlags = bind,
                CpuAccessFlags = CpuAccessFlags.Write,
                OptionFlags = ResourceOptionFlags.BufferStructured,
                StructureByteStride = Utilities.SizeOf<T>()
            });
        }

        private Buffer CreateDefaultStructuredBuffer<T>(int elementCount, BindFlags bind) where T : struct
        {
            return new Buffer(_device, new BufferDescription
            {
                SizeInBytes = Utilities.SizeOf<T>() * elementCount,
                Usage = ResourceUsage.Default,
                BindFlags = bind,
                CpuAccessFlags = CpuAccessFlags.None,
                OptionFlags = ResourceOptionFlags.BufferStructured,
                StructureByteStride = Utilities.SizeOf<T>()
            });
        }

        private void UpdateDynamicStructured<T>(Buffer buffer, T[] data) where T : struct
        {
            DataStream ds;
            _ctx.MapSubresource(buffer, 0, MapMode.WriteDiscard, MapFlags.None, out ds);
            ds.WriteRange(data, 0, data.Length);
            _ctx.UnmapSubresource(buffer, 0);
            ds.Dispose();
        }

        private void UpdateConstants(Constants c)
        {
            DataStream ds;
            _ctx.MapSubresource(_cbConstants, 0, MapMode.WriteDiscard, MapFlags.None, out ds);
            ds.Write(c);
            _ctx.UnmapSubresource(_cbConstants, 0);
            ds.Dispose();
        }

        private static int NextPow2(int v)
        {
            int p = 1;
            while (p < v) p <<= 1;
            return p;
        }

        private static TimeSpan ToTimeSpan(long ticks)
        {
            if (ticks <= 0)
            {
                return TimeSpan.Zero;
            }

            return TimeSpan.FromSeconds(ticks / (double)Stopwatch.Frequency);
        }

        private void DisposeTriangles()
        {
            _triSrv?.Dispose();
            _triSrv = null;
            _triBuffer?.Dispose();
            _triBuffer = null;
            _triCapacity = 0;
        }

        private void DisposePoints()
        {
            _ptSrv?.Dispose();
            _ptSrv = null;
            _ptBuffer?.Dispose();
            _ptBuffer = null;
            _ptCapacity = 0;
        }

        private void DisposePointZones()
        {
            _pointZoneSrv?.Dispose();
            _pointZoneSrv = null;
            _pointZoneBuffer?.Dispose();
            _pointZoneBuffer = null;
            _pointZoneCapacity = 0;
        }

        private void DisposeZoneRanges()
        {
            _zoneRangesSrv?.Dispose();
            _zoneRangesSrv = null;
            _zoneRangesBuffer?.Dispose();
            _zoneRangesBuffer = null;
            _zoneRangesCapacity = 0;
        }

        private void DisposeOutput()
        {
            _outUav?.Dispose();
            _outUav = null;
            _outBuffer?.Dispose();
            _outBuffer = null;
            _outStaging?.Dispose();
            _outStaging = null;
            _outCapacity = 0;
        }

        public void Dispose()
        {
            DisposeOutput();
            DisposeZoneRanges();
            DisposePointZones();
            DisposePoints();
            DisposeTriangles();

            _cbConstants?.Dispose();
            _cbConstants = null;
            _cs?.Dispose();
            _ctx?.Dispose();
            _device?.Dispose();
        }

        // Single shader source, dual-ray optional (UseSecondRay constant)
        private const string ShaderSource = @"
struct Triangle
{
    float4 v0;
    float4 v1;
    float4 v2;
};

struct ZoneRange
{
    uint triStart;
    uint triCount;
};

cbuffer Constants : register(b0)
{
    uint TriangleCount;
    uint PointCount;
    uint ZoneCount;
    uint UseSecondRay;
};

StructuredBuffer<Triangle> Triangles : register(t0);
StructuredBuffer<float4> Points : register(t1);
StructuredBuffer<uint> PointZones : register(t2);
StructuredBuffer<ZoneRange> ZoneRanges : register(t3);
RWStructuredBuffer<uint> InsideOut : register(u0);

float3 Jitter(uint idx, float3 dirBase)
{
    // deterministic tiny jitter to reduce edge/vertex degeneracy
    uint s = idx * 1664525u + 1013904223u;
    float j = ((s & 1023u) / 1023.0f) * 2.0f - 1.0f;
    float3 d = normalize(dirBase + float3(j * 1e-3, j * 0.37e-3, j * 0.51e-3));
    return d;
}

bool RayTri(float3 orig, float3 dir, Triangle t)
{
    float3 v0 = t.v0.xyz;
    float3 v1 = t.v1.xyz;
    float3 v2 = t.v2.xyz;

    float3 e1 = v1 - v0;
    float3 e2 = v2 - v0;

    float3 p = cross(dir, e2);
    float det = dot(e1, p);

    if (abs(det) < 1e-8) return false;

    float invDet = 1.0 / det;

    float3 tv = orig - v0;
    float u = dot(tv, p) * invDet;
    if (u < 0.0 || u > 1.0) return false;

    float3 q = cross(tv, e1);
    float v = dot(dir, q) * invDet;
    if (v < 0.0 || (u + v) > 1.0) return false;

    float dist = dot(e2, q) * invDet;
    return dist > 1e-5;
}

uint ParityRange(float3 p, float3 dir, ZoneRange r)
{
    uint hits = 0u;
    uint end = r.triStart + r.triCount;
    [loop]
    for (uint i = r.triStart; i < end; i++)
    {
        if (RayTri(p, dir, Triangles[i])) hits++;
    }
    return (hits & 1u);
}

[numthreads(256, 1, 1)]
void CSMain(uint3 tid : SV_DispatchThreadID)
{
    uint idx = tid.x;
    if (idx >= PointCount) return;

    uint zid = PointZones[idx];
    if (zid >= ZoneCount)
    {
        InsideOut[idx] = 0u;
        return;
    }

    ZoneRange zr = ZoneRanges[zid];
    float3 p = Points[idx].xyz;

    float3 d1 = Jitter(idx, float3(0.976, 0.182, 0.120));
    uint inside1 = ParityRange(p, d1, zr);

    if (UseSecondRay == 0u)
    {
        InsideOut[idx] = inside1;
        return;
    }

    float3 d2 = Jitter(idx, float3(0.289, 0.957, 0.034));
    uint inside2 = ParityRange(p, d2, zr);

    InsideOut[idx] = (inside1 == inside2) ? inside1 : 2u;
}";
    }
}
