@echo off
setlocal
cd /d "%~dp0"

REM Compile application resources.
rc.exe SpectrogramContextMenu.rc
if errorlevel 1 (
    echo [ERROR] Resource compilation failed
    exit /b 1
)

REM Build one per-user installer/runtime executable. FFmpeg is not bundled.
cl.exe /nologo /EHsc /O2 /W4 /utf-8 /DUNICODE /D_UNICODE /D_WIN64 ^
    SpectrogramApp.cpp SpectrogramContextMenu.cpp SpectrogramContextMenu.res ^
    /link /MACHINE:X64 /SUBSYSTEM:WINDOWS /MANIFEST:NO ^
    shlwapi.lib shell32.lib advapi32.lib user32.lib gdi32.lib gdiplus.lib ^
    /OUT:SpectrogramContextMenu.exe
if errorlevel 1 (
    echo [ERROR] Build failed
    exit /b 1
)

echo [OK] SpectrogramContextMenu.exe created
echo Run the EXE to install for the current user.
exit /b 0
