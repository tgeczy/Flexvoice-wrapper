cmake_minimum_required(VERSION 3.20)
project(flexvoice_wrap LANGUAGES CXX)

add_library(fvwrap SHARED
  src/fvwrap.cpp
)

target_include_directories(fvwrap PRIVATE
  ${CMAKE_CURRENT_SOURCE_DIR}/include
  ${CMAKE_CURRENT_SOURCE_DIR}/vendor/flexvoice_sdk/include
)

target_compile_definitions(fvwrap PRIVATE
  FVWRAP_EXPORTS
  WIN32_LEAN_AND_MEAN
  NOMINMAX
)

set_target_properties(fvwrap PROPERTIES
  CXX_STANDARD 14
  CXX_STANDARD_REQUIRED YES
)

if (MSVC)
  target_compile_options(fvwrap PRIVATE /EHsc /GR /W3)
endif()

# Pick the right lib folder for x86 vs x64:
if (CMAKE_SIZEOF_VOID_P EQUAL 8)
  target_link_directories(fvwrap PRIVATE ${CMAKE_CURRENT_SOURCE_DIR}/vendor/flexvoice_sdk/lib/x64)
else()
  target_link_directories(fvwrap PRIVATE ${CMAKE_CURRENT_SOURCE_DIR}/vendor/flexvoice_sdk/lib/x86)
endif()

# Link whatever your engine needs. Start with the engine + farm server libs;
# add FVNetClient/FVNetServer only if the linker asks for missing symbols.
target_link_libraries(fvwrap PRIVATE
  FlexVoice_3_01_001.lib
  FVFarmServer_3_01_001.lib
  FVNetClient.lib
  FVNetServer.lib
)
