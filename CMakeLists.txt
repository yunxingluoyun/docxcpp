cmake_minimum_required(VERSION 3.20)

project(python_docx_cpp LANGUAGES CXX)

find_package(Python3 REQUIRED COMPONENTS Interpreter)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

add_library(
  docxcpp
  src/document.cpp
  src/document_paragraph_format.cpp
  src/header_footer.cpp
  src/hyperlink_comment.cpp
  src/layout.cpp
  src/media.cpp
  src/opc_package.cpp
  src/paragraph.cpp
  src/paragraph_ops.cpp
  src/section.cpp
  src/table.cpp
  3rdparty/pugixml-1.13/pugixml.cpp
)

target_include_directories(
  docxcpp
  PUBLIC
    ${CMAKE_CURRENT_SOURCE_DIR}/include
    ${CMAKE_CURRENT_SOURCE_DIR}/3rdparty/pugixml-1.13
  PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/3rdparty/miniz-3.0.2
)

target_compile_definitions(
  docxcpp
  PRIVATE
    DOCXCPP_DEFAULT_TEMPLATE_DIR="${CMAKE_CURRENT_SOURCE_DIR}/src/docx/templates/default-docx-template"
)

add_executable(docxcpp_example examples/minimal.cpp)
target_link_libraries(docxcpp_example PRIVATE docxcpp)

add_executable(docxcpp_full_feature_example examples/full_feature_5_pages.cpp)
target_link_libraries(docxcpp_full_feature_example PRIVATE docxcpp)

enable_testing()

add_executable(docxcpp_smoke_test tests/smoke_test.cpp)
target_link_libraries(docxcpp_smoke_test PRIVATE docxcpp)
target_compile_definitions(
  docxcpp_smoke_test
  PRIVATE
    DOCXCPP_SOURCE_DIR="${CMAKE_CURRENT_SOURCE_DIR}"
    DOCXCPP_TEST_ARTIFACT_DIR="${CMAKE_CURRENT_BINARY_DIR}/test-artifacts"
)

add_test(NAME docxcpp_smoke_test COMMAND docxcpp_smoke_test)

add_executable(docxcpp_api_coverage_test tests/api_coverage_test.cpp)
target_link_libraries(docxcpp_api_coverage_test PRIVATE docxcpp)
target_compile_definitions(
  docxcpp_api_coverage_test
  PRIVATE
    DOCXCPP_SOURCE_DIR="${CMAKE_CURRENT_SOURCE_DIR}"
    DOCXCPP_TEST_ARTIFACT_DIR="${CMAKE_CURRENT_BINARY_DIR}/test-artifacts"
)

add_test(NAME docxcpp_api_coverage_test COMMAND docxcpp_api_coverage_test)

add_test(
  NAME docxcpp_full_feature_example_generate
  COMMAND docxcpp_full_feature_example ${CMAKE_CURRENT_BINARY_DIR}/test-artifacts/full-feature-5-pages.docx
)
set_tests_properties(
  docxcpp_full_feature_example_generate
  PROPERTIES
    WORKING_DIRECTORY ${CMAKE_CURRENT_SOURCE_DIR}
)

add_test(
  NAME docxcpp_python_docx_regression
  COMMAND ${Python3_EXECUTABLE} ${CMAKE_CURRENT_SOURCE_DIR}/tests/python_docx_regression.py
          ${CMAKE_CURRENT_BINARY_DIR}/test-artifacts
)
set_tests_properties(
  docxcpp_python_docx_regression
  PROPERTIES
    DEPENDS "docxcpp_smoke_test;docxcpp_api_coverage_test;docxcpp_full_feature_example_generate"
)
