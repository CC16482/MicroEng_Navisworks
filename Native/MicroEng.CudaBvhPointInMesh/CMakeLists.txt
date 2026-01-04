cmake_minimum_required(VERSION 3.24)
project(MicroEngCudaBvhPointInMesh LANGUAGES CXX CUDA)

add_library(MicroEng.CudaBvhPointInMesh SHARED microeng_cuda_bvh_point_in_mesh.cu)
set_target_properties(MicroEng.CudaBvhPointInMesh PROPERTIES
    CUDA_SEPARABLE_COMPILATION ON
    CXX_STANDARD 17
    CUDA_ARCHITECTURES native
)

# Strongly recommended for deployment: avoid requiring cudart64_*.dll beside your plugin
set_target_properties(MicroEng.CudaBvhPointInMesh PROPERTIES CUDA_RUNTIME_LIBRARY Static)
