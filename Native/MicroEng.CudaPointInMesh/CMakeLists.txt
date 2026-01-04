cmake_minimum_required(VERSION 3.24)
project(MicroEngCudaPointInMesh LANGUAGES CXX CUDA)

add_library(MicroEng.CudaPointInMesh SHARED microeng_cuda_point_in_mesh.cu)
set_target_properties(MicroEng.CudaPointInMesh PROPERTIES
    CUDA_SEPARABLE_COMPILATION ON
    CXX_STANDARD 17
    CUDA_ARCHITECTURES native
)

# Optional: static link the CUDA runtime so you don't need cudart64_*.dll deployed.
# Uncomment if you want that behavior.
# set_target_properties(MicroEng.CudaPointInMesh PROPERTIES CUDA_RUNTIME_LIBRARY Static)
