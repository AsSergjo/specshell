cmake_minimum_required(VERSION 3.10)
project(SpectrogramContextMenu)

set(CMAKE_CXX_STANDARD 11)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# Определяем исходные файлы
set(SOURCES
    SpectrogramApp.cpp
    SpectrogramContextMenu.cpp
)

set(RESOURCES
    SpectrogramContextMenu.rc
)

set(HEADERS
    SpectrogramContextMenu.h
)

# Standalone per-user installer and runtime application.
add_executable(SpectrogramContextMenu WIN32
    ${SOURCES}
    ${HEADERS}
    ${RESOURCES}
)

# Определяем UNICODE
target_compile_definitions(SpectrogramContextMenu PRIVATE
    UNICODE
    _UNICODE
)

# Линкуем необходимые библиотеки
target_link_libraries(SpectrogramContextMenu
    shlwapi
    shell32
    advapi32
    user32
    gdi32
    gdiplus
)

# Настройки для Visual Studio
if(MSVC)
    target_compile_options(SpectrogramContextMenu PRIVATE
        /W4
        /O2
        /EHsc
        /utf-8
    )
    set_property(TARGET SpectrogramContextMenu APPEND_STRING PROPERTY
        LINK_FLAGS " /MANIFEST:NO")
endif()

# Настройки для MinGW
if(MINGW)
    target_compile_options(SpectrogramContextMenu PRIVATE
        -O2
        -static-libgcc
        -static-libstdc++
    )
endif()

# Устанавливаем выходное имя
set_target_properties(SpectrogramContextMenu PROPERTIES
    OUTPUT_NAME "SpectrogramContextMenu"
    PREFIX ""
)

# Информация о проекте
message(STATUS "===========================================")
message(STATUS "Spectrogram Context Menu application")
message(STATUS "===========================================")
message(STATUS "Компилятор: ${CMAKE_CXX_COMPILER_ID}")
message(STATUS "C++ стандарт: ${CMAKE_CXX_STANDARD}")
message(STATUS "Тип сборки: ${CMAKE_BUILD_TYPE}")
message(STATUS "===========================================")
