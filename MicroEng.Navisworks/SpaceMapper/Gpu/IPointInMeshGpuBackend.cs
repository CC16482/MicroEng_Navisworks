using System;
using System.Threading;

namespace MicroEng.Navisworks.SpaceMapper.Gpu
{
    internal interface IPointInMeshGpuBackend : IDisposable
    {
        string BackendName { get; }
        string DeviceName { get; }

        /// <summary>
        /// Returns flags per point: 0=outside, 1=inside, 2=uncertain (needs CPU fallback).
        /// </summary>
        uint[] TestPoints(Triangle[] trianglesLocal, Float4[] pointsLocal, bool intensive, CancellationToken ct);
    }
}
