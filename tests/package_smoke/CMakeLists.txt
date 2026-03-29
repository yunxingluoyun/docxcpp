cmake_minimum_required(VERSION 3.20)

project(docxcpp_package_smoke LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

find_package(docxcpp CONFIG REQUIRED)

add_executable(docxcpp_package_smoke main.cpp)
target_link_libraries(docxcpp_package_smoke PRIVATE docxcpp::docxcpp)
