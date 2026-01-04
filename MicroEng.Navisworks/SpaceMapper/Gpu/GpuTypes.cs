using System.Runtime.InteropServices;

namespace MicroEng.Navisworks.SpaceMapper.Gpu
{
    [StructLayout(LayoutKind.Sequential)]
    internal struct Float4
    {
        public float X, Y, Z, W;

        public Float4(float x, float y, float z, float w = 0f)
        {
            X = x;
            Y = y;
            Z = z;
            W = w;
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct Triangle
    {
        public Float4 V0;
        public Float4 V1;
        public Float4 V2;

        public Triangle(Float4 v0, Float4 v1, Float4 v2)
        {
            V0 = v0;
            V1 = v1;
            V2 = v2;
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct ZoneRange
    {
        public uint TriStart;
        public uint TriCount;
    }
}
