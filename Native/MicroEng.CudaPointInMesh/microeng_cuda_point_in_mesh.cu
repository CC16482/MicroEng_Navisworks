// Native/MicroEng.CudaPointInMesh/microeng_cuda_point_in_mesh.cu
#include <cuda_runtime.h>
#include <stdint.h>
#include <string.h>
#include <math.h>

struct Float4 { float x, y, z, w; };
struct Triangle { Float4 v0, v1, v2; };
struct ZoneRange { uint32_t triStart; uint32_t triCount; };

static int g_device = -1;
static Triangle* d_tris = nullptr;
static Float4* d_pts = nullptr;
static uint32_t* d_pointZones = nullptr;
static ZoneRange* d_zoneRanges = nullptr;
static uint32_t* d_out = nullptr;
static int g_triCap = 0;
static int g_ptCap = 0;
static int g_zoneCap = 0;

static void writeErr(char* buf, int len, const char* msg)
{
    if (!buf || len <= 0) return;
    strncpy(buf, msg ? msg : "", len - 1);
    buf[len - 1] = '\0';
}

static int ensureCap(int triCount, int ptCount, char* err, int errLen)
{
    if (triCount > g_triCap)
    {
        if (d_tris) cudaFree(d_tris);
        g_triCap = triCount;
        cudaError_t e = cudaMalloc((void**)&d_tris, sizeof(Triangle) * (size_t)g_triCap);
        if (e != cudaSuccess) { writeErr(err, errLen, cudaGetErrorString(e)); return -1; }
    }
    if (ptCount > g_ptCap)
    {
        if (d_pts) cudaFree(d_pts);
        if (d_pointZones) cudaFree(d_pointZones);
        if (d_out) cudaFree(d_out);
        g_ptCap = ptCount;
        cudaError_t e1 = cudaMalloc((void**)&d_pts, sizeof(Float4) * (size_t)g_ptCap);
        if (e1 != cudaSuccess) { writeErr(err, errLen, cudaGetErrorString(e1)); return -2; }
        cudaError_t e1b = cudaMalloc((void**)&d_pointZones, sizeof(uint32_t) * (size_t)g_ptCap);
        if (e1b != cudaSuccess) { writeErr(err, errLen, cudaGetErrorString(e1b)); return -2; }
        cudaError_t e2 = cudaMalloc((void**)&d_out, sizeof(uint32_t) * (size_t)g_ptCap);
        if (e2 != cudaSuccess) { writeErr(err, errLen, cudaGetErrorString(e2)); return -3; }
    }
    return 0;
}

static int ensureZoneCap(int zoneCount, char* err, int errLen)
{
    if (zoneCount > g_zoneCap)
    {
        if (d_zoneRanges) cudaFree(d_zoneRanges);
        g_zoneCap = zoneCount;
        cudaError_t e = cudaMalloc((void**)&d_zoneRanges, sizeof(ZoneRange) * (size_t)g_zoneCap);
        if (e != cudaSuccess) { writeErr(err, errLen, cudaGetErrorString(e)); return -1; }
    }
    return 0;
}

__device__ __forceinline__ float3 make3(const Float4& f) { return make_float3(f.x, f.y, f.z); }
__device__ __forceinline__ float3 sub3(const float3& a, const float3& b) { return make_float3(a.x-b.x, a.y-b.y, a.z-b.z); }
__device__ __forceinline__ float3 cross3(const float3& a, const float3& b)
{
    return make_float3(a.y*b.z - a.z*b.y, a.z*b.x - a.x*b.z, a.x*b.y - a.y*b.x);
}
__device__ __forceinline__ float dot3(const float3& a, const float3& b) { return a.x*b.x + a.y*b.y + a.z*b.z; }

__device__ __forceinline__ float3 norm3(const float3& v)
{
    float d = sqrtf(dot3(v, v));
    if (d <= 1e-20f) return make_float3(0,0,0);
    return make_float3(v.x/d, v.y/d, v.z/d);
}

__device__ __forceinline__ bool rayTri(const float3& orig, const float3& dir, const Triangle& t)
{
    // Match the HLSL tolerances as closely as possible
    const float EPS_DET = 1e-8f;
    const float EPS_DIST = 1e-5f;

    float3 v0 = make3(t.v0);
    float3 v1 = make3(t.v1);
    float3 v2 = make3(t.v2);

    float3 e1 = sub3(v1, v0);
    float3 e2 = sub3(v2, v0);
    float3 pvec = cross3(dir, e2);
    float det = dot3(e1, pvec);

    if (fabsf(det) < EPS_DET) return false;
    float invDet = 1.0f / det;

    float3 tvec = sub3(orig, v0);
    float u = dot3(tvec, pvec) * invDet;
    if (u < 0.f || u > 1.f) return false;

    float3 qvec = cross3(tvec, e1);
    float v = dot3(dir, qvec) * invDet;
    if (v < 0.f || (u + v) > 1.f) return false;

    float dist = dot3(e2, qvec) * invDet;
    return dist > EPS_DIST;
}

__device__ __forceinline__ bool insideRay(const float3& p, const Triangle* tris, int triCount, const float3& baseDir, uint32_t tid)
{
    // Deterministic jitter (to mimic HLSL intent: avoid edge/vertex degeneracy)
    uint32_t s = tid * 1664525u + 1013904223u;
    float jitter = ((s & 1023u) / 1023.0f) * 2.0f - 1.0f;
    float3 dir = norm3(make_float3(
        baseDir.x + jitter * 1e-3f,
        baseDir.y + jitter * 0.37e-3f,
        baseDir.z + jitter * 0.51e-3f));

    int hits = 0;
    for (int i = 0; i < triCount; ++i)
        if (rayTri(p, dir, tris[i])) hits++;

    return (hits & 1) != 0;
}

__device__ __forceinline__ bool insideRayRange(const float3& p, const Triangle* tris, int triStart, int triCount, const float3& baseDir, uint32_t tid)
{
    uint32_t s = tid * 1664525u + 1013904223u;
    float jitter = ((s & 1023u) / 1023.0f) * 2.0f - 1.0f;
    float3 dir = norm3(make_float3(
        baseDir.x + jitter * 1e-3f,
        baseDir.y + jitter * 0.37e-3f,
        baseDir.z + jitter * 0.51e-3f));

    int hits = 0;
    for (int i = 0; i < triCount; ++i)
        if (rayTri(p, dir, tris[triStart + i])) hits++;

    return (hits & 1) != 0;
}

__global__ void pointInMeshKernel(const Triangle* tris, int triCount, const Float4* pts, int ptCount, int useSecondRay, uint32_t* outFlags)
{
    int tid = (int)(blockIdx.x * blockDim.x + threadIdx.x);
    if (tid >= ptCount) return;

    float3 p = make_float3(pts[tid].x, pts[tid].y, pts[tid].z);

    // Match HLSL-ish rays
    const float3 base1 = make_float3(0.976f, 0.182f, 0.120f);
    const float3 base2 = make_float3(0.289f, 0.957f, 0.034f);

    bool in1 = insideRay(p, tris, triCount, base1, (uint32_t)tid);

    if (!useSecondRay)
    {
        outFlags[tid] = in1 ? 1u : 0u;
        return;
    }

    bool in2 = insideRay(p, tris, triCount, base2, (uint32_t)tid);

    outFlags[tid] = (in1 == in2) ? (in1 ? 1u : 0u) : 2u;
}

__global__ void pointInMeshKernelBatched(
    const Triangle* tris,
    int triCount,
    const Float4* pts,
    int ptCount,
    const uint32_t* pointZones,
    const ZoneRange* zoneRanges,
    int zoneCount,
    int useSecondRay,
    uint32_t* outFlags)
{
    int tid = (int)(blockIdx.x * blockDim.x + threadIdx.x);
    if (tid >= ptCount) return;

    if (zoneCount <= 0)
    {
        outFlags[tid] = 0;
        return;
    }

    uint32_t zid = pointZones[tid];
    if (zid >= (uint32_t)zoneCount)
    {
        outFlags[tid] = 0;
        return;
    }

    ZoneRange zr = zoneRanges[zid];
    int triStart = (int)zr.triStart;
    int triCountLocal = (int)zr.triCount;
    if (triCountLocal <= 0 || triStart < 0 || triStart + triCountLocal > triCount)
    {
        outFlags[tid] = 0;
        return;
    }

    float3 p = make_float3(pts[tid].x, pts[tid].y, pts[tid].z);

    const float3 base1 = make_float3(0.976f, 0.182f, 0.120f);
    const float3 base2 = make_float3(0.289f, 0.957f, 0.034f);

    bool in1 = insideRayRange(p, tris, triStart, triCountLocal, base1, (uint32_t)tid);

    if (!useSecondRay)
    {
        outFlags[tid] = in1 ? 1u : 0u;
        return;
    }

    bool in2 = insideRayRange(p, tris, triStart, triCountLocal, base2, (uint32_t)tid);
    outFlags[tid] = (in1 == in2) ? (in1 ? 1u : 0u) : 2u;
}

extern "C" __declspec(dllexport)
int me_cuda_init(int deviceOrdinal, int* chosenDevice, char* errBuf, int errBufLen)
{
    int count = 0;
    cudaError_t e = cudaGetDeviceCount(&count);
    if (e != cudaSuccess || count <= 0) { writeErr(errBuf, errBufLen, "No CUDA device found."); return -1; }

    int best = 0;
    size_t bestMem = 0;
    for (int i = 0; i < count; ++i)
    {
        cudaDeviceProp p{};
        if (cudaGetDeviceProperties(&p, i) == cudaSuccess)
        {
            if ((size_t)p.totalGlobalMem > bestMem)
            {
                bestMem = (size_t)p.totalGlobalMem;
                best = i;
            }
        }
    }

    int dev = (deviceOrdinal >= 0 && deviceOrdinal < count) ? deviceOrdinal : best;

    e = cudaSetDevice(dev);
    if (e != cudaSuccess) { writeErr(errBuf, errBufLen, cudaGetErrorString(e)); return -2; }

    g_device = dev;
    if (chosenDevice) *chosenDevice = dev;
    return 0;
}

extern "C" __declspec(dllexport)
int me_cuda_test_points(const Triangle* triangles, int triCount,
                        const Float4* points, int ptCount,
                        int useSecondRay,
                        uint32_t* outFlags,
                        char* errBuf, int errBufLen)
{
    if (!triangles || triCount <= 0 || !points || ptCount <= 0 || !outFlags)
    { writeErr(errBuf, errBufLen, "Invalid arguments."); return -1; }

    int capRc = ensureCap(triCount, ptCount, errBuf, errBufLen);
    if (capRc != 0) return capRc;

    cudaError_t e1 = cudaMemcpy(d_tris, triangles, sizeof(Triangle) * (size_t)triCount, cudaMemcpyHostToDevice);
    if (e1 != cudaSuccess) { writeErr(errBuf, errBufLen, cudaGetErrorString(e1)); return -2; }

    cudaError_t e2 = cudaMemcpy(d_pts, points, sizeof(Float4) * (size_t)ptCount, cudaMemcpyHostToDevice);
    if (e2 != cudaSuccess) { writeErr(errBuf, errBufLen, cudaGetErrorString(e2)); return -3; }

    int threads = 256;
    int blocks = (ptCount + threads - 1) / threads;
    pointInMeshKernel<<<blocks, threads>>>(d_tris, triCount, d_pts, ptCount, useSecondRay ? 1 : 0, d_out);

    cudaError_t e3 = cudaGetLastError();
    if (e3 != cudaSuccess) { writeErr(errBuf, errBufLen, cudaGetErrorString(e3)); return -4; }

    cudaError_t e4 = cudaDeviceSynchronize();
    if (e4 != cudaSuccess) { writeErr(errBuf, errBufLen, cudaGetErrorString(e4)); return -5; }

    cudaError_t e5 = cudaMemcpy(outFlags, d_out, sizeof(uint32_t) * (size_t)ptCount, cudaMemcpyDeviceToHost);
    if (e5 != cudaSuccess) { writeErr(errBuf, errBufLen, cudaGetErrorString(e5)); return -6; }

    return 0;
}

extern "C" __declspec(dllexport)
int me_cuda_test_points_batched(const Triangle* triangles, int triCount,
                                const Float4* points, int ptCount,
                                const uint32_t* pointZones,
                                const ZoneRange* zoneRanges, int zoneCount,
                                int useSecondRay,
                                uint32_t* outFlags,
                                char* errBuf, int errBufLen)
{
    if (!triangles || triCount <= 0 || !points || ptCount <= 0 || !outFlags || !pointZones || !zoneRanges || zoneCount <= 0)
    { writeErr(errBuf, errBufLen, "Invalid arguments."); return -1; }

    int capRc = ensureCap(triCount, ptCount, errBuf, errBufLen);
    if (capRc != 0) return capRc;

    int zoneRc = ensureZoneCap(zoneCount, errBuf, errBufLen);
    if (zoneRc != 0) return zoneRc;

    cudaError_t e1 = cudaMemcpy(d_tris, triangles, sizeof(Triangle) * (size_t)triCount, cudaMemcpyHostToDevice);
    if (e1 != cudaSuccess) { writeErr(errBuf, errBufLen, cudaGetErrorString(e1)); return -2; }

    cudaError_t e2 = cudaMemcpy(d_pts, points, sizeof(Float4) * (size_t)ptCount, cudaMemcpyHostToDevice);
    if (e2 != cudaSuccess) { writeErr(errBuf, errBufLen, cudaGetErrorString(e2)); return -3; }

    cudaError_t e2b = cudaMemcpy(d_pointZones, pointZones, sizeof(uint32_t) * (size_t)ptCount, cudaMemcpyHostToDevice);
    if (e2b != cudaSuccess) { writeErr(errBuf, errBufLen, cudaGetErrorString(e2b)); return -3; }

    cudaError_t e2c = cudaMemcpy(d_zoneRanges, zoneRanges, sizeof(ZoneRange) * (size_t)zoneCount, cudaMemcpyHostToDevice);
    if (e2c != cudaSuccess) { writeErr(errBuf, errBufLen, cudaGetErrorString(e2c)); return -3; }

    int threads = 256;
    int blocks = (ptCount + threads - 1) / threads;
    pointInMeshKernelBatched<<<blocks, threads>>>(d_tris, triCount, d_pts, ptCount, d_pointZones, d_zoneRanges, zoneCount, useSecondRay ? 1 : 0, d_out);

    cudaError_t e3 = cudaGetLastError();
    if (e3 != cudaSuccess) { writeErr(errBuf, errBufLen, cudaGetErrorString(e3)); return -4; }

    cudaError_t e4 = cudaDeviceSynchronize();
    if (e4 != cudaSuccess) { writeErr(errBuf, errBufLen, cudaGetErrorString(e4)); return -5; }

    cudaError_t e5 = cudaMemcpy(outFlags, d_out, sizeof(uint32_t) * (size_t)ptCount, cudaMemcpyDeviceToHost);
    if (e5 != cudaSuccess) { writeErr(errBuf, errBufLen, cudaGetErrorString(e5)); return -6; }

    return 0;
}

extern "C" __declspec(dllexport)
void me_cuda_shutdown()
{
    if (d_tris) cudaFree(d_tris);
    if (d_pts) cudaFree(d_pts);
    if (d_pointZones) cudaFree(d_pointZones);
    if (d_zoneRanges) cudaFree(d_zoneRanges);
    if (d_out) cudaFree(d_out);
    d_tris = nullptr; d_pts = nullptr; d_pointZones = nullptr; d_zoneRanges = nullptr; d_out = nullptr;
    g_triCap = 0; g_ptCap = 0; g_zoneCap = 0;
    g_device = -1;
}
