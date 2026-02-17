cmake_minimum_required(VERSION 3.10)
project(SpectrogramContextMenu)

set(CMAKE_CXX_STANDARD 11)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# Определяем исходные файлы
set(SOURCES
    SpectrogramContextMenu.cpp
    Register.cpp
)

set(RESOURCES
    SpectrogramContextMenu.rc
)

set(HEADERS
    SpectrogramContextMenu.h
)

# Создаем DLL
add_library(SpectrogramContextMenu SHARED
    ${SOURCES}
    ${HEADERS}
    ${RESOURCES}
    SpectrogramContextMenu.def
)

# Определяем UNICODE
target_compile_definitions(SpectrogramContextMenu PRIVATE
    UNICODE
    _UNICODE
)

# Линкуем необходимые библиотеки
target_link_libraries(SpectrogramContextMenu
    ole32
    oleaut32
    uuid
    shlwapi
    shell32
    advapi32
    gdi32
    comctl32
)

# Настройки для Visual Studio
if(MSVC)
    target_compile_options(SpectrogramContextMenu PRIVATE
        /W3
        /O2
        /EHsc
    )
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

# Команды для установки и удаления
if(WIN32)
    # Создаем custom targets для регистрации
    add_custom_target(install_extension
        COMMAND regsvr32 /s "$<TARGET_FILE:SpectrogramContextMenu>"
        DEPENDS SpectrogramContextMenu
        COMMENT "Регистрация Shell Extension (требуются права администратора)"
    )

    add_custom_target(uninstall_extension
        COMMAND regsvr32 /s /u "$<TARGET_FILE:SpectrogramContextMenu>"
        COMMENT "Отмена регистрации Shell Extension (требуются права администратора)"
    )
endif()

# Информация о проекте
message(STATUS "===========================================")
message(STATUS "Spectrogram Context Menu Shell Extension")
message(STATUS "===========================================")
message(STATUS "Компилятор: ${CMAKE_CXX_COMPILER_ID}")
message(STATUS "C++ стандарт: ${CMAKE_CXX_STANDARD}")
message(STATUS "Тип сборки: ${CMAKE_BUILD_TYPE}")
message(STATUS "===========================================")
