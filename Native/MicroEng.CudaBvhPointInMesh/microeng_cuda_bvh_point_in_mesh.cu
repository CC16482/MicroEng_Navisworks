#include <cuda_runtime.h>
#include <stdint.h>
#include <vector>
#include <unordered_map>
#include <algorithm>
#include <cmath>
#include <cstring>

struct Float4 { float x, y, z, w; };
struct Triangle { Float4 v0, v1, v2; };
struct ZoneRange { uint32_t triStart; uint32_t triCount; };

// BVH node (48 bytes, aligned)
struct BvhNode
{
    Float4 bmin;
    Float4 bmax;
    int left;
    int right;
    int triStart;
    int triCount; // >0 => leaf
};

static void writeErr(char* buf, int len, const char* msg)
{
    if (!buf || len <= 0) return;
    std::strncpy(buf, msg ? msg : "", len - 1);
    buf[len - 1] = '\0';
}

static inline Float4 make4(float x, float y, float z, float w = 0.f) { Float4 f{ x, y, z, w }; return f; }
static inline void triBounds(const Triangle& t, Float4& mn, Float4& mx)
{
    float x0 = t.v0.x, y0 = t.v0.y, z0 = t.v0.z;
    float x1 = t.v1.x, y1 = t.v1.y, z1 = t.v1.z;
    float x2 = t.v2.x, y2 = t.v2.y, z2 = t.v2.z;
    mn = make4(fminf(x0, fminf(x1, x2)), fminf(y0, fminf(y1, y2)), fminf(z0, fminf(z1, z2)), 0.f);
    mx = make4(fmaxf(x0, fmaxf(x1, x2)), fmaxf(y0, fmaxf(y1, y2)), fmaxf(z0, fmaxf(z1, z2)), 0.f);
}
static inline Float4 triCentroid(const Triangle& t)
{
    return make4((t.v0.x + t.v1.x + t.v2.x) / 3.f, (t.v0.y + t.v1.y + t.v2.y) / 3.f, (t.v0.z + t.v1.z + t.v2.z) / 3.f, 0.f);
}

struct Scene
{
    Triangle* d_tris = nullptr;
    int triCount = 0;

    BvhNode* d_nodes = nullptr;
    int nodeCount = 0;

    int* d_zoneRoots = nullptr;
    int zoneCount = 0;

    // reusable point buffers
    Float4* d_points = nullptr;
    uint32_t* d_pointZones = nullptr;
    uint32_t* d_out = nullptr;
    int ptCap = 0;

    ~Scene()
    {
        if (d_tris) cudaFree(d_tris);
        if (d_nodes) cudaFree(d_nodes);
        if (d_zoneRoots) cudaFree(d_zoneRoots);
        if (d_points) cudaFree(d_points);
        if (d_pointZones) cudaFree(d_pointZones);
        if (d_out) cudaFree(d_out);
    }
};

static std::unordered_map<int, Scene*> g_scenes;
static int g_nextHandle = 1;

static inline Float4 bmin4(const Float4& a, const Float4& b) { return make4(fminf(a.x, b.x), fminf(a.y, b.y), fminf(a.z, b.z), 0.f); }
static inline Float4 bmax4(const Float4& a, const Float4& b) { return make4(fmaxf(a.x, b.x), fmaxf(a.y, b.y), fmaxf(a.z, b.z), 0.f); }

static void computeRangeBounds(const std::vector<Triangle>& tris, int start, int count, Float4& outMin, Float4& outMax)
{
    Float4 mn = make4(1e30f, 1e30f, 1e30f, 0);
    Float4 mx = make4(-1e30f, -1e30f, -1e30f, 0);
    for (int i = 0; i < count; i++)
    {
        Float4 tmin, tmax; triBounds(tris[start + i], tmin, tmax);
        mn = bmin4(mn, tmin);
        mx = bmax4(mx, tmax);
    }
    outMin = mn; outMax = mx;
}

static int buildBvhRecursive(std::vector<Triangle>& tris, int start, int count, int leafSize, std::vector<BvhNode>& nodes)
{
    BvhNode n{};
    Float4 mn, mx;
    computeRangeBounds(tris, start, count, mn, mx);
    n.bmin = mn;
    n.bmax = mx;

    if (count <= leafSize)
    {
        n.left = -1; n.right = -1;
        n.triStart = start;
        n.triCount = count;
        int idx = (int)nodes.size();
        nodes.push_back(n);
        return idx;
    }

    // centroid bounds
    Float4 cmn = make4(1e30f, 1e30f, 1e30f, 0);
    Float4 cmx = make4(-1e30f, -1e30f, -1e30f, 0);
    for (int i = 0; i < count; i++)
    {
        Float4 c = triCentroid(tris[start + i]);
        cmn = bmin4(cmn, c);
        cmx = bmax4(cmx, c);
    }
    float ex = cmx.x - cmn.x;
    float ey = cmx.y - cmn.y;
    float ez = cmx.z - cmn.z;
    int axis = (ex >= ey && ex >= ez) ? 0 : (ey >= ez ? 1 : 2);

    int mid = start + count / 2;

    auto key = [axis](const Triangle& t)->float
    {
        Float4 c = triCentroid(t);
        return axis == 0 ? c.x : (axis == 1 ? c.y : c.z);
    };

    std::nth_element(tris.begin() + start, tris.begin() + mid, tris.begin() + start + count,
        [&](const Triangle& a, const Triangle& b) { return key(a) < key(b); });

    int idx = (int)nodes.size();
    nodes.push_back(n); // placeholder

    int left = buildBvhRecursive(tris, start, mid - start, leafSize, nodes);
    int right = buildBvhRecursive(tris, mid, start + count - mid, leafSize, nodes);

    nodes[idx].left = left;
    nodes[idx].right = right;
    nodes[idx].triStart = 0;
    nodes[idx].triCount = 0;
    return idx;
}

__device__ __forceinline__ float3 make3(const Float4& f) { return make_float3(f.x, f.y, f.z); }
__device__ __forceinline__ float3 sub3(const float3& a, const float3& b) { return make_float3(a.x - b.x, a.y - b.y, a.z - b.z); }
__device__ __forceinline__ float3 cross3(const float3& a, const float3& b)
{
    return make_float3(a.y * b.z - a.z * b.y, a.z * b.x - a.x * b.z, a.x * b.y - a.y * b.x);
}
__device__ __forceinline__ float dot3(const float3& a, const float3& b) { return a.x * b.x + a.y * b.y + a.z * b.z; }

__device__ __forceinline__ bool rayAabb(const float3& o, const float3& invD, const Float4& mn4, const Float4& mx4)
{
    float3 mn = make_float3(mn4.x, mn4.y, mn4.z);
    float3 mx = make_float3(mx4.x, mx4.y, mx4.z);

    float t1 = (mn.x - o.x) * invD.x;
    float t2 = (mx.x - o.x) * invD.x;
    float tmin = fminf(t1, t2);
    float tmax = fmaxf(t1, t2);

    t1 = (mn.y - o.y) * invD.y;
    t2 = (mx.y - o.y) * invD.y;
    tmin = fmaxf(tmin, fminf(t1, t2));
    tmax = fminf(tmax, fmaxf(t1, t2));

    t1 = (mn.z - o.z) * invD.z;
    t2 = (mx.z - o.z) * invD.z;
    tmin = fmaxf(tmin, fminf(t1, t2));
    tmax = fminf(tmax, fmaxf(t1, t2));

    return tmax > fmaxf(tmin, 1e-5f);
}

__device__ __forceinline__ bool rayTri(const float3& orig, const float3& dir, const Triangle& t)
{
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
    float invDet = 1.f / det;

    float3 tvec = sub3(orig, v0);
    float u = dot3(tvec, pvec) * invDet;
    if (u < 0.f || u > 1.f) return false;

    float3 qvec = cross3(tvec, e1);
    float v = dot3(dir, qvec) * invDet;
    if (v < 0.f || (u + v) > 1.f) return false;

    float dist = dot3(e2, qvec) * invDet;
    return dist > EPS_DIST;
}

__device__ __forceinline__ float3 jitterDir(uint32_t idx, float3 base)
{
    uint32_t s = idx * 1664525u + 1013904223u;
    float j = ((s & 1023u) / 1023.0f) * 2.0f - 1.0f;
    float3 d = make_float3(base.x + j * 1e-3f, base.y + j * 0.37e-3f, base.z + j * 0.51e-3f);
    float len = sqrtf(dot3(d, d));
    return (len > 1e-20f) ? make_float3(d.x / len, d.y / len, d.z / len) : base;
}

__device__ __forceinline__ uint32_t parityBvh(const float3& p, const float3& dir, const Triangle* tris, const BvhNode* nodes, int root)
{
    float3 invD = make_float3(1.f / dir.x, 1.f / dir.y, 1.f / dir.z);
    int stack[64];
    int sp = 0;
    stack[sp++] = root;

    int hits = 0;

    while (sp > 0)
    {
        int ni = stack[--sp];
        const BvhNode& n = nodes[ni];
        if (!rayAabb(p, invD, n.bmin, n.bmax)) continue;

        if (n.triCount > 0)
        {
            for (int i = 0; i < n.triCount; i++)
            {
                if (rayTri(p, dir, tris[n.triStart + i])) hits++;
            }
        }
        else
        {
            // push children
            if (sp < 62)
            {
                stack[sp++] = n.left;
                stack[sp++] = n.right;
            }
        }
    }

    return (hits & 1) ? 1u : 0u;
}

__global__ void pointInMeshBvhKernel(
    const Triangle* tris,
    const BvhNode* nodes,
    const int* zoneRoots,
    int zoneCount,
    const Float4* points,
    const uint32_t* pointZones,
    int pointCount,
    int useSecondRay,
    uint32_t* outFlags)
{
    int idx = (int)(blockIdx.x * blockDim.x + threadIdx.x);
    if (idx >= pointCount) return;

    uint32_t zid = pointZones[idx];
    if ((int)zid >= zoneCount) { outFlags[idx] = 0; return; }

    int root = zoneRoots[zid];
    float3 p = make_float3(points[idx].x, points[idx].y, points[idx].z);

    float3 d1 = jitterDir((uint32_t)idx, make_float3(0.976f, 0.182f, 0.120f));
    uint32_t in1 = parityBvh(p, d1, tris, nodes, root);

    if (!useSecondRay)
    {
        outFlags[idx] = in1;
        return;
    }

    float3 d2 = jitterDir((uint32_t)idx, make_float3(0.289f, 0.957f, 0.034f));
    uint32_t in2 = parityBvh(p, d2, tris, nodes, root);

    outFlags[idx] = (in1 == in2) ? in1 : 2u;
}

extern "C" __declspec(dllexport)
int me_cuda_bvh_create_scene(
    const Triangle* trianglesAll, int triTotal,
    const ZoneRange* zones, int zoneCount,
    int leafSize,
    int* outHandle,
    char* errBuf, int errLen)
{
    if (!trianglesAll || triTotal <= 0 || !zones || zoneCount <= 0 || !outHandle)
    {
        writeErr(errBuf, errLen, "Invalid arguments.");
        return -1;
    }

    try
    {
        // copy triangles so we can reorder per-zone during BVH build
        std::vector<Triangle> tris(trianglesAll, trianglesAll + triTotal);

        std::vector<BvhNode> allNodes;
        allNodes.reserve(zoneCount * 128);

        std::vector<int> zoneRoots(zoneCount, -1);

        for (int z = 0; z < zoneCount; z++)
        {
            const ZoneRange zr = zones[z];
            if (zr.triCount == 0) { zoneRoots[z] = -1; continue; }

            int start = (int)zr.triStart;
            int count = (int)zr.triCount;
            if (start < 0 || start + count > triTotal)
            {
                writeErr(errBuf, errLen, "ZoneRange out of bounds.");
                return -2;
            }

            // build BVH for this zone (in-place reorder within range)
            std::vector<BvhNode> localNodes;
            localNodes.reserve(count * 2);

            int rootLocal = buildBvhRecursive(tris, start, count, leafSize, localNodes);

            // append nodes to allNodes with index fixups (children indices)
            int nodeBase = (int)allNodes.size();
            for (size_t i = 0; i < localNodes.size(); i++)
            {
                BvhNode n = localNodes[i];

                if (n.triCount == 0)
                {
                    n.left += nodeBase;
                    n.right += nodeBase;
                }
                // leaf triStart already absolute (start in global tri array)
                allNodes.push_back(n);
            }

            zoneRoots[z] = nodeBase + rootLocal;
        }

        Scene* sc = new Scene();
        sc->triCount = triTotal;
        sc->nodeCount = (int)allNodes.size();
        sc->zoneCount = zoneCount;

        cudaError_t e;

        e = cudaMalloc((void**)&sc->d_tris, sizeof(Triangle) * (size_t)sc->triCount);
        if (e != cudaSuccess) { writeErr(errBuf, errLen, cudaGetErrorString(e)); delete sc; return -10; }

        e = cudaMemcpy(sc->d_tris, tris.data(), sizeof(Triangle) * (size_t)sc->triCount, cudaMemcpyHostToDevice);
        if (e != cudaSuccess) { writeErr(errBuf, errLen, cudaGetErrorString(e)); delete sc; return -11; }

        e = cudaMalloc((void**)&sc->d_nodes, sizeof(BvhNode) * (size_t)sc->nodeCount);
        if (e != cudaSuccess) { writeErr(errBuf, errLen, cudaGetErrorString(e)); delete sc; return -12; }

        e = cudaMemcpy(sc->d_nodes, allNodes.data(), sizeof(BvhNode) * (size_t)sc->nodeCount, cudaMemcpyHostToDevice);
        if (e != cudaSuccess) { writeErr(errBuf, errLen, cudaGetErrorString(e)); delete sc; return -13; }

        e = cudaMalloc((void**)&sc->d_zoneRoots, sizeof(int) * (size_t)sc->zoneCount);
        if (e != cudaSuccess) { writeErr(errBuf, errLen, cudaGetErrorString(e)); delete sc; return -14; }

        e = cudaMemcpy(sc->d_zoneRoots, zoneRoots.data(), sizeof(int) * (size_t)sc->zoneCount, cudaMemcpyHostToDevice);
        if (e != cudaSuccess) { writeErr(errBuf, errLen, cudaGetErrorString(e)); delete sc; return -15; }

        int handle = g_nextHandle++;
        g_scenes[handle] = sc;
        *outHandle = handle;
        return 0;
    }
    catch (...)
    {
        writeErr(errBuf, errLen, "Exception in create_scene.");
        return -99;
    }
}

static int ensurePointCap(Scene* sc, int ptCount, char* errBuf, int errLen)
{
    if (ptCount <= sc->ptCap) return 0;
    if (sc->d_points) cudaFree(sc->d_points);
    if (sc->d_pointZones) cudaFree(sc->d_pointZones);
    if (sc->d_out) cudaFree(sc->d_out);

    sc->ptCap = ptCount;
    cudaError_t e;

    e = cudaMalloc((void**)&sc->d_points, sizeof(Float4) * (size_t)sc->ptCap);
    if (e != cudaSuccess) { writeErr(errBuf, errLen, cudaGetErrorString(e)); return -1; }

    e = cudaMalloc((void**)&sc->d_pointZones, sizeof(uint32_t) * (size_t)sc->ptCap);
    if (e != cudaSuccess) { writeErr(errBuf, errLen, cudaGetErrorString(e)); return -2; }

    e = cudaMalloc((void**)&sc->d_out, sizeof(uint32_t) * (size_t)sc->ptCap);
    if (e != cudaSuccess) { writeErr(errBuf, errLen, cudaGetErrorString(e)); return -3; }

    return 0;
}

extern "C" __declspec(dllexport)
int me_cuda_bvh_test_points(
    int sceneHandle,
    const Float4* points, const uint32_t* pointZones, int pointCount,
    int useSecondRay,
    uint32_t* outFlags,
    char* errBuf, int errLen)
{
    auto it = g_scenes.find(sceneHandle);
    if (it == g_scenes.end()) { writeErr(errBuf, errLen, "Invalid scene handle."); return -1; }
    Scene* sc = it->second;

    if (!points || !pointZones || pointCount <= 0 || !outFlags)
    {
        writeErr(errBuf, errLen, "Invalid arguments.");
        return -2;
    }

    int capRc = ensurePointCap(sc, pointCount, errBuf, errLen);
    if (capRc != 0) return -10 + capRc;

    cudaError_t e;
    e = cudaMemcpy(sc->d_points, points, sizeof(Float4) * (size_t)pointCount, cudaMemcpyHostToDevice);
    if (e != cudaSuccess) { writeErr(errBuf, errLen, cudaGetErrorString(e)); return -20; }

    e = cudaMemcpy(sc->d_pointZones, pointZones, sizeof(uint32_t) * (size_t)pointCount, cudaMemcpyHostToDevice);
    if (e != cudaSuccess) { writeErr(errBuf, errLen, cudaGetErrorString(e)); return -21; }

    int threads = 256;
    int blocks = (pointCount + threads - 1) / threads;
    pointInMeshBvhKernel<<<blocks, threads>>>(
        sc->d_tris, sc->d_nodes, sc->d_zoneRoots, sc->zoneCount,
        sc->d_points, sc->d_pointZones, pointCount,
        useSecondRay ? 1 : 0,
        sc->d_out);

    e = cudaGetLastError();
    if (e != cudaSuccess) { writeErr(errBuf, errLen, cudaGetErrorString(e)); return -30; }

    e = cudaDeviceSynchronize();
    if (e != cudaSuccess) { writeErr(errBuf, errLen, cudaGetErrorString(e)); return -31; }

    e = cudaMemcpy(outFlags, sc->d_out, sizeof(uint32_t) * (size_t)pointCount, cudaMemcpyDeviceToHost);
    if (e != cudaSuccess) { writeErr(errBuf, errLen, cudaGetErrorString(e)); return -32; }

    return 0;
}

extern "C" __declspec(dllexport)
void me_cuda_bvh_destroy_scene(int sceneHandle)
{
    auto it = g_scenes.find(sceneHandle);
    if (it == g_scenes.end()) return;
    delete it->second;
    g_scenes.erase(it);
}
