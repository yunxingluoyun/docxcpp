cmake_minimum_required(VERSION 3.20)

project(python_docx_cpp LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

add_library(
  docxcpp
  cpp/src/document.cpp
  cpp/src/opc_package.cpp
  3rdparty/pugixml-1.13/pugixml.cpp
)

target_include_directories(
  docxcpp
  PUBLIC
    ${CMAKE_CURRENT_SOURCE_DIR}/cpp/include
    ${CMAKE_CURRENT_SOURCE_DIR}/3rdparty/pugixml-1.13
  PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/3rdparty/miniz-3.0.2
)

target_compile_definitions(
  docxcpp
  PRIVATE
    DOCXCPP_DEFAULT_TEMPLATE_DIR="${CMAKE_CURRENT_SOURCE_DIR}/src/docx/templates/default-docx-template"
)

add_executable(docxcpp_example cpp/examples/minimal.cpp)
target_link_libraries(docxcpp_example PRIVATE docxcpp)

enable_testing()

add_executable(docxcpp_smoke_test cpp/tests/smoke_test.cpp)
target_link_libraries(docxcpp_smoke_test PRIVATE docxcpp)
target_compile_definitions(
  docxcpp_smoke_test
  PRIVATE
    DOCXCPP_SOURCE_DIR="${CMAKE_CURRENT_SOURCE_DIR}"
    DOCXCPP_TEST_ARTIFACT_DIR="${CMAKE_CURRENT_BINARY_DIR}/test-artifacts"
)

add_test(NAME docxcpp_smoke_test COMMAND docxcpp_smoke_test)

add_executable(docxcpp_api_coverage_test cpp/tests/api_coverage_test.cpp)
target_link_libraries(docxcpp_api_coverage_test PRIVATE docxcpp)
target_compile_definitions(
  docxcpp_api_coverage_test
  PRIVATE
    DOCXCPP_SOURCE_DIR="${CMAKE_CURRENT_SOURCE_DIR}"
    DOCXCPP_TEST_ARTIFACT_DIR="${CMAKE_CURRENT_BINARY_DIR}/test-artifacts"
)

add_test(NAME docxcpp_api_coverage_test COMMAND docxcpp_api_coverage_test)
