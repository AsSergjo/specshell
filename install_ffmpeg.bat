@echo off
chcp 1251
echo ========================================
echo FFmpeg Installation Helper
echo ========================================
echo.

REM Проверяем, установлен ли уже FFmpeg
where ffmpeg.exe >nul 2>&1
if %errorlevel% equ 0 (
    echo FFmpeg уже установлен и доступен в PATH!
    ffmpeg -version | findstr "ffmpeg version"
    echo.
    pause
    exit /b 0
)

echo FFmpeg не найден в системе.
echo.
echo Этот скрипт поможет установить FFmpeg.
echo.
echo Выберите способ установки:
echo.
echo 1. Скачать и установить автоматически (рекомендуется)
echo 2. Показать инструкцию для ручной установки
echo 3. Выход
echo.

choice /c 123 /n /m "Ваш выбор: "

if errorlevel 3 exit /b 0
if errorlevel 2 goto manual
if errorlevel 1 goto automatic

:automatic
echo.
echo Автоматическая установка...
echo.

REM Проверяем наличие curl или powershell
where curl.exe >nul 2>&1
if %errorlevel% equ 0 (
    set DOWNLOADER=curl
) else (
    set DOWNLOADER=powershell
)

echo Создаем директорию для FFmpeg...
if not exist "C:\ffmpeg\bin\" mkdir "C:\ffmpeg\bin\"

echo.
echo Скачиваем FFmpeg (это может занять несколько минут)...
echo Используется: %DOWNLOADER%
echo.

if "%DOWNLOADER%"=="curl" (
    curl -L "https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-essentials.zip" -o "%TEMP%\ffmpeg.zip"
) else (
    powershell -Command "& {[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-essentials.zip' -OutFile '%TEMP%\ffmpeg.zip'}"
)

if not exist "%TEMP%\ffmpeg.zip" (
    echo.
    echo Ошибка: Не удалось скачать FFmpeg.
    echo Попробуйте ручную установку (вариант 2).
    echo.
    pause
    exit /b 1
)

echo.
echo Распаковка архива...
powershell -Command "Expand-Archive -Path '%TEMP%\ffmpeg.zip' -DestinationPath '%TEMP%\ffmpeg_temp' -Force"

echo.
echo Копирование файлов...
for /d %%i in ("%TEMP%\ffmpeg_temp\ffmpeg-*") do (
    xcopy "%%i\bin\ffmpeg.exe" "C:\ffmpeg\bin\" /Y /I
    xcopy "%%i\bin\ffprobe.exe" "C:\ffmpeg\bin\" /Y /I
)

echo.
echo Очистка временных файлов...
del "%TEMP%\ffmpeg.zip" >nul 2>&1
rmdir /s /q "%TEMP%\ffmpeg_temp" >nul 2>&1

REM Добавляем в PATH
echo.
echo Добавление FFmpeg в PATH...
setx PATH "%PATH%;C:\ffmpeg\bin" >nul 2>&1

echo.
echo ========================================
echo Установка завершена успешно!
echo ========================================
echo.
echo FFmpeg установлен в: C:\ffmpeg\bin\
echo.
echo ВАЖНО: Закройте и откройте заново командную строку,
echo чтобы изменения PATH вступили в силу.
echo.
pause
exit /b 0

:manual
echo.
echo ========================================
echo Ручная установка FFmpeg
echo ========================================
echo.
echo Шаг 1: Скачайте FFmpeg
echo    Перейдите на: https://www.gyan.dev/ffmpeg/builds/
echo    Скачайте: ffmpeg-release-essentials.zip
echo.
echo Шаг 2: Распакуйте архив
echo    Распакуйте содержимое в папку, например: C:\ffmpeg
echo.
echo Шаг 3: Добавьте в PATH
echo    1. Нажмите Win + X и выберите "Система"
echo    2. Нажмите "Дополнительные параметры системы"
echo    3. Нажмите "Переменные среды"
echo    4. В разделе "Системные переменные" найдите "Path"
echo    5. Нажмите "Изменить"
echo    6. Нажмите "Создать"
echo    7. Добавьте: C:\ffmpeg\bin
echo    8. Нажмите "ОК" везде
echo.
echo Шаг 4: Проверьте установку
echo    Откройте новую командную строку и выполните:
echo    ffmpeg -version
echo.
echo Альтернативные источники:
echo    - https://ffmpeg.org/download.html
echo    - https://github.com/BtbN/FFmpeg-Builds/releases
echo.
pause
exit /b 0
