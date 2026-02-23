@echo off
REM Переходим в директорию, где расположен этот bat-файл
cd /d "%~dp0"

REM Очистка предыдущей сборки (опционально)
REM if exist build rmdir /s /q build

REM Создаем папку build, если её нет
if not exist build mkdir build

REM Конфигурация CMake с генератором Visual Studio 2022
cmake -S . -B build -G "Visual Studio 17 2022"

REM Сборка в конфигурации Release
cmake --build build --config Release

REM Сообщение об успехе
if %ERRORLEVEL% EQU 0 (
    echo OK: build\Release\loader.exe
) else (
    echo ERROR: %ERRORLEVEL%
)

pause