using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace MicroEng.Navisworks.SpaceMapper.Gpu
{
    internal sealed class CudaBvhPointInMeshGpu : IDisposable
    {
        private const string DllName = "MicroEng.CudaBvhPointInMesh.dll";

        public int SceneHandle { get; private set; }

        public static bool TryCreateScene(
            Triangle[] trianglesAll,
            ZoneRange[] zoneRanges,
            int leafSize,
            out CudaBvhPointInMeshGpu scene,
            out string reason)
        {
            scene = null;
            reason = null;

            if (trianglesAll == null)
            {
                reason = "trianglesAll is null.";
                return false;
            }

            if (zoneRanges == null)
            {
                reason = "zoneRanges is null.";
                return false;
            }

            if (trianglesAll.Length == 0 || zoneRanges.Length == 0)
            {
                reason = "No triangles or zones to build a scene.";
                return false;
            }

            try
            {
                LoadNativeFromPluginFolder(DllName);

                var err = new StringBuilder(2048);
                int handle;
                int rc = me_cuda_bvh_create_scene(
                    trianglesAll,
                    trianglesAll.Length,
                    zoneRanges,
                    zoneRanges.Length,
                    leafSize,
                    out handle,
                    err,
                    err.Capacity);
                if (rc != 0)
                {
                    reason = $"CUDA BVH create_scene failed ({rc}): {err}";
                    return false;
                }

                scene = new CudaBvhPointInMeshGpu { SceneHandle = handle };
                return true;
            }
            catch (Exception ex)
            {
                reason = ex.Message;
                return false;
            }
        }

        public uint[] TestPoints(
            Float4[] pointsAll,
            uint[] pointZoneIds,
            bool intensiveTwoRays,
            CancellationToken token)
        {
            if (SceneHandle == 0) throw new InvalidOperationException("SceneHandle is not initialized.");
            if (pointsAll == null) throw new ArgumentNullException(nameof(pointsAll));
            if (pointZoneIds == null) throw new ArgumentNullException(nameof(pointZoneIds));
            if (pointsAll.Length != pointZoneIds.Length)
            {
                throw new ArgumentException("pointZoneIds length must match pointsAll length.", nameof(pointZoneIds));
            }

            if (pointsAll.Length == 0)
            {
                return Array.Empty<uint>();
            }

            token.ThrowIfCancellationRequested();

            var result = new uint[pointsAll.Length];
            var err = new StringBuilder(2048);
            int rc = me_cuda_bvh_test_points(
                SceneHandle,
                pointsAll,
                pointZoneIds,
                pointsAll.Length,
                intensiveTwoRays ? 1 : 0,
                result,
                err,
                err.Capacity);
            if (rc != 0)
            {
                throw new InvalidOperationException($"CUDA BVH test_points failed ({rc}): {err}");
            }

            return result;
        }

        public void Dispose()
        {
            if (SceneHandle != 0)
            {
                me_cuda_bvh_destroy_scene(SceneHandle);
                SceneHandle = 0;
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
        private static extern int me_cuda_bvh_create_scene(
            [In] Triangle[] trianglesAll, int triTotal,
            [In] ZoneRange[] zones, int zoneCount,
            int leafSize,
            out int outHandle,
            StringBuilder errBuf, int errLen);

        [DllImport(DllName, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern int me_cuda_bvh_test_points(
            int sceneHandle,
            [In] Float4[] points,
            [In] uint[] pointZones,
            int pointCount,
            int useSecondRay,
            [Out] uint[] outFlags,
            StringBuilder errBuf, int errLen);

        [DllImport(DllName, CallingConvention = CallingConvention.Cdecl)]
        private static extern void me_cuda_bvh_destroy_scene(int sceneHandle);
    }
}
