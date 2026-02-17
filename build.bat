@echo off
REM Компилируем ресурсный файл
rc.exe SpectrogramContextMenu.rc
if %errorlevel% neq 0 (
    echo [ERROR] Resource compilation failed
    pause
    exit /b 1
)

REM Компилируем исходный код с включением ресурсного файла
cl.exe /LD /EHsc /O2 /W3 /D "UNICODE" /D "_UNICODE" /D "_WIN64" SpectrogramContextMenu.cpp Register.cpp SpectrogramContextMenu.res /link /MACHINE:X64 /DEF:SpectrogramContextMenu.def ole32.lib oleaut32.lib uuid.lib shlwapi.lib shell32.lib advapi32.lib user32.lib gdi32.lib comctl32.lib /OUT:SpectrogramContextMenu.dll
if %errorlevel% equ 0 (
    echo.
    echo [OK] Build successful
    echo.
    echo Next steps:
    echo 1. regsvr32 SpectrogramContextMenu.dll
    echo 2. Restart explorer.exe
) else (
    echo [ERROR] Build failed
)
pause
