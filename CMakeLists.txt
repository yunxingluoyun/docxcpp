cmake_minimum_required(VERSION 3.20)

project(docxcpp VERSION 0.1.0 LANGUAGES CXX)

include(GNUInstallDirs)
include(CMakePackageConfigHelpers)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)
set(CMAKE_CXX_VISIBILITY_PRESET hidden)
set(CMAKE_VISIBILITY_INLINES_HIDDEN YES)

set(DOCXCPP_DEFAULT_TEMPLATE_ARCHIVE
    "${CMAKE_CURRENT_SOURCE_DIR}/src/docx/templates/default.docx"
)
set(DOCXCPP_GENERATED_DIR "${CMAKE_CURRENT_BINARY_DIR}/generated")
set(DOCXCPP_GENERATED_TEMPLATE_CPP "${DOCXCPP_GENERATED_DIR}/default_template.cpp")
set(DOCXCPP_PACKAGE_CONFIG_DIR "${CMAKE_CURRENT_BINARY_DIR}/cmake")

add_custom_command(
  OUTPUT ${DOCXCPP_GENERATED_TEMPLATE_CPP}
  COMMAND ${CMAKE_COMMAND} -E make_directory ${DOCXCPP_GENERATED_DIR}
  COMMAND ${CMAKE_COMMAND}
          -DINPUT=${DOCXCPP_DEFAULT_TEMPLATE_ARCHIVE}
          -DOUTPUT=${DOCXCPP_GENERATED_TEMPLATE_CPP}
          -P ${CMAKE_CURRENT_SOURCE_DIR}/cmake/GenerateDefaultTemplate.cmake
  DEPENDS
    ${CMAKE_CURRENT_SOURCE_DIR}/cmake/GenerateDefaultTemplate.cmake
    ${DOCXCPP_DEFAULT_TEMPLATE_ARCHIVE}
  VERBATIM
)

add_library(
  docxcpp
  ${DOCXCPP_GENERATED_TEMPLATE_CPP}
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

add_library(docxcpp::docxcpp ALIAS docxcpp)

target_include_directories(
  docxcpp
  PUBLIC
    $<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/include>
    $<INSTALL_INTERFACE:${CMAKE_INSTALL_INCLUDEDIR}>
  PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/src
    ${CMAKE_CURRENT_SOURCE_DIR}/3rdparty/pugixml-1.13
    ${CMAKE_CURRENT_SOURCE_DIR}/3rdparty/miniz-3.0.2
)

target_compile_definitions(
  docxcpp
  PUBLIC
    $<$<NOT:$<BOOL:${BUILD_SHARED_LIBS}>>:DOCXCPP_STATIC_DEFINE>
  PRIVATE
    DOCXCPP_BUILDING_LIBRARY
)

set_target_properties(
  docxcpp
  PROPERTIES
    VERSION ${PROJECT_VERSION}
    SOVERSION ${PROJECT_VERSION_MAJOR}
    EXPORT_NAME docxcpp
)

add_executable(docxcpp_example examples/minimal.cpp)
target_link_libraries(docxcpp_example PRIVATE docxcpp)

add_executable(docxcpp_full_feature_example examples/full_feature_5_pages.cpp)
target_link_libraries(docxcpp_full_feature_example PRIVATE docxcpp)

add_executable(docxcpp_read_existing_example examples/read_existing.cpp)
target_link_libraries(docxcpp_read_existing_example PRIVATE docxcpp)

add_executable(docxcpp_sections_headers_footers_example examples/sections_headers_footers.cpp)
target_link_libraries(docxcpp_sections_headers_footers_example PRIVATE docxcpp)

add_executable(docxcpp_styles_and_tables_example examples/styles_and_tables.cpp)
target_link_libraries(docxcpp_styles_and_tables_example PRIVATE docxcpp)

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
target_sources(docxcpp_api_coverage_test PRIVATE 3rdparty/pugixml-1.13/pugixml.cpp)
target_include_directories(
  docxcpp_api_coverage_test
  PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/3rdparty/pugixml-1.13
)
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
  NAME docxcpp_package_install_smoke
  COMMAND
    ${CMAKE_COMMAND}
    -DDOCXCPP_SOURCE_DIR=${CMAKE_CURRENT_SOURCE_DIR}
    -DDOCXCPP_BINARY_DIR=${CMAKE_CURRENT_BINARY_DIR}
    -DDOCXCPP_BUILD_CONFIG=$<CONFIG>
    -P ${CMAKE_CURRENT_SOURCE_DIR}/cmake/VerifyInstalledPackage.cmake
)
set_tests_properties(
  docxcpp_package_install_smoke
  PROPERTIES
    DEPENDS "docxcpp_smoke_test;docxcpp_api_coverage_test"
)

configure_package_config_file(
  ${CMAKE_CURRENT_SOURCE_DIR}/cmake/docxcppConfig.cmake.in
  ${DOCXCPP_PACKAGE_CONFIG_DIR}/docxcppConfig.cmake
  INSTALL_DESTINATION ${CMAKE_INSTALL_LIBDIR}/cmake/docxcpp
)

write_basic_package_version_file(
  ${DOCXCPP_PACKAGE_CONFIG_DIR}/docxcppConfigVersion.cmake
  VERSION ${PROJECT_VERSION}
  COMPATIBILITY SameMajorVersion
)

install(
  TARGETS docxcpp
  EXPORT docxcppTargets
  ARCHIVE DESTINATION ${CMAKE_INSTALL_LIBDIR}
  LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR}
  RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR}
  INCLUDES DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}
)

install(
  DIRECTORY ${CMAKE_CURRENT_SOURCE_DIR}/include/docxcpp
  DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}
)

install(
  EXPORT docxcppTargets
  FILE docxcppTargets.cmake
  NAMESPACE docxcpp::
  DESTINATION ${CMAKE_INSTALL_LIBDIR}/cmake/docxcpp
)

install(
  FILES
    ${DOCXCPP_PACKAGE_CONFIG_DIR}/docxcppConfig.cmake
    ${DOCXCPP_PACKAGE_CONFIG_DIR}/docxcppConfigVersion.cmake
  DESTINATION ${CMAKE_INSTALL_LIBDIR}/cmake/docxcpp
)
