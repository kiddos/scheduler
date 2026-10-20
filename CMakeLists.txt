cmake_minimum_required(VERSION 3.18)
project(scheduler LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTOUIC ON)
set(CMAKE_AUTORCC ON)

find_package(Qt5 COMPONENTS Widgets Sql REQUIRED)

set(BUILD_DEPS ON CACHE BOOL "Build all dependencies" FORCE)

include(FetchContent)
FetchContent_Declare(
  ortools
  GIT_REPOSITORY https://github.com/google/or-tools.git
  GIT_TAG v9.8
)
FetchContent_MakeAvailable(ortools)

add_executable(scheduler main.cpp scheduler.cpp scheduler.h scheduler.ui)
target_link_libraries(scheduler PRIVATE ortools::ortools Qt5::Widgets Qt5::Sql)
