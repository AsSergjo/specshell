@echo off
REM Unregisters the context menu handler.
REM This script must be run as an administrator.

REM The name of the DLL to unregister.
set DLL_NAME=SpectrogramContextMenu.dll

echo Unregistering %DLL_NAME%...
regsvr32 /u %DLL_NAME%

if %errorlevel% equ 0 (
    echo.
    echo Successfully unregistered the extension.
    echo.
    echo To apply the changes, Windows Explorer needs to be restarted.
    echo This will cause your desktop and taskbar to disappear for a moment.
    echo.
    choice /c YN /m "Do you want to restart Explorer now?"

    if %errorlevel% equ 1 (
        echo Restarting Explorer...
        taskkill /f /im explorer.exe >nul
        start explorer.exe
    )

    echo.
    echo Uninstallation complete.
) else (
    echo.
    echo Failed to unregister the extension. Make sure you are running this script as an Administrator.
)

pause
