using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace MicroEng.Navisworks.SpaceMapper.Gpu
{
    internal sealed class CudaPointInMeshGpu : IPointInMeshGpuBackend
    {
        private const string DllName = "MicroEng.CudaPointInMesh.dll";

        public string BackendName => "CUDA";
        public string DeviceName { get; private set; } = "CUDA device";

        private bool _initialized;

        private CudaPointInMeshGpu()
        {
        }

        public static bool TryCreate(out CudaPointInMeshGpu backend, out string reason)
        {
            backend = null;
            reason = null;

            try
            {
                // Ensure the native DLL is loadable from the plugin folder.
                LoadNativeFromPluginFolder(DllName);

                var err = new StringBuilder(2048);
                int chosen;
                int rc = me_cuda_init(-1, out chosen, err, err.Capacity);
                if (rc != 0)
                {
                    reason = $"me_cuda_init failed ({rc}): {err}";
                    return false;
                }

                backend = new CudaPointInMeshGpu
                {
                    _initialized = true,
                    DeviceName = $"CUDA device {chosen}"
                };
                return true;
            }
            catch (Exception ex)
            {
                reason = ex.Message;
                return false;
            }
        }

        public uint[] TestPoints(Triangle[] trianglesLocal, Float4[] pointsLocal, bool intensive, CancellationToken ct)
        {
            if (!_initialized) throw new InvalidOperationException("CUDA backend not initialized.");
            ct.ThrowIfCancellationRequested();

            var outFlags = new uint[pointsLocal.Length];
            var err = new StringBuilder(2048);

            GCHandle hTris = default;
            GCHandle hPts = default;
            GCHandle hOut = default;
            try
            {
                hTris = GCHandle.Alloc(trianglesLocal, GCHandleType.Pinned);
                hPts = GCHandle.Alloc(pointsLocal, GCHandleType.Pinned);
                hOut = GCHandle.Alloc(outFlags, GCHandleType.Pinned);

                int rc = me_cuda_test_points(
                    hTris.AddrOfPinnedObject(), trianglesLocal.Length,
                    hPts.AddrOfPinnedObject(), pointsLocal.Length,
                    intensive ? 1 : 0,
                    hOut.AddrOfPinnedObject(),
                    err, err.Capacity);

                if (rc != 0)
                {
                    throw new InvalidOperationException($"me_cuda_test_points failed ({rc}): {err}");
                }

                return outFlags;
            }
            finally
            {
                if (hTris.IsAllocated) hTris.Free();
                if (hPts.IsAllocated) hPts.Free();
                if (hOut.IsAllocated) hOut.Free();
            }
        }

        public uint[] TestPointsBatched(
            Triangle[] trianglesAll,
            Float4[] pointsAll,
            uint[] pointZoneIds,
            ZoneRange[] zoneRanges,
            bool intensive,
            CancellationToken ct)
        {
            if (!_initialized) throw new InvalidOperationException("CUDA backend not initialized.");
            ct.ThrowIfCancellationRequested();

            if (pointsAll == null) throw new ArgumentNullException(nameof(pointsAll));
            if (pointZoneIds == null) throw new ArgumentNullException(nameof(pointZoneIds));
            if (zoneRanges == null) throw new ArgumentNullException(nameof(zoneRanges));
            if (pointsAll.Length != pointZoneIds.Length)
            {
                throw new ArgumentException("pointZoneIds must match pointsAll length.", nameof(pointZoneIds));
            }

            if (pointsAll.Length == 0)
            {
                return Array.Empty<uint>();
            }

            var outFlags = new uint[pointsAll.Length];
            var err = new StringBuilder(2048);

            GCHandle hTris = default;
            GCHandle hPts = default;
            GCHandle hZones = default;
            GCHandle hRanges = default;
            GCHandle hOut = default;
            try
            {
                hTris = GCHandle.Alloc(trianglesAll, GCHandleType.Pinned);
                hPts = GCHandle.Alloc(pointsAll, GCHandleType.Pinned);
                hZones = GCHandle.Alloc(pointZoneIds, GCHandleType.Pinned);
                hRanges = GCHandle.Alloc(zoneRanges, GCHandleType.Pinned);
                hOut = GCHandle.Alloc(outFlags, GCHandleType.Pinned);

                int rc = me_cuda_test_points_batched(
                    hTris.AddrOfPinnedObject(), trianglesAll.Length,
                    hPts.AddrOfPinnedObject(), pointsAll.Length,
                    hZones.AddrOfPinnedObject(),
                    hRanges.AddrOfPinnedObject(), zoneRanges.Length,
                    intensive ? 1 : 0,
                    hOut.AddrOfPinnedObject(),
                    err, err.Capacity);

                if (rc != 0)
                {
                    throw new InvalidOperationException($"me_cuda_test_points_batched failed ({rc}): {err}");
                }

                return outFlags;
            }
            finally
            {
                if (hTris.IsAllocated) hTris.Free();
                if (hPts.IsAllocated) hPts.Free();
                if (hZones.IsAllocated) hZones.Free();
                if (hRanges.IsAllocated) hRanges.Free();
                if (hOut.IsAllocated) hOut.Free();
            }
        }

        public void Dispose()
        {
            if (_initialized)
            {
                me_cuda_shutdown();
                _initialized = false;
            }
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr LoadLibrary(string lpFileName);

        private static void LoadNativeFromPluginFolder(string dllName)
        {
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            var path = System.IO.Path.Combine(baseDir, dllName);
            if (System.IO.File.Exists(path))
            {
                LoadLibrary(path);
            }
        }

        [DllImport(DllName, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern int me_cuda_init(int deviceOrdinal, out int chosenDevice, StringBuilder errBuf, int errBufLen);

        [DllImport(DllName, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern int me_cuda_test_points(
            IntPtr triangles, int triCount,
            IntPtr points, int ptCount,
            int useSecondRay,
            IntPtr outFlags,
            StringBuilder errBuf, int errBufLen);

        [DllImport(DllName, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern int me_cuda_test_points_batched(
            IntPtr triangles, int triCount,
            IntPtr points, int ptCount,
            IntPtr pointZones,
            IntPtr zoneRanges, int zoneCount,
            int useSecondRay,
            IntPtr outFlags,
            StringBuilder errBuf, int errBufLen);

        [DllImport(DllName, CallingConvention = CallingConvention.Cdecl)]
        private static extern void me_cuda_shutdown();
    }
}
