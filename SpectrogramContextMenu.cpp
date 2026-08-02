#include "SpectrogramContextMenu.h"

#include <windows.h>
#include <gdiplus.h>
#include <shellapi.h>
#include <shlwapi.h>

#include <cmath>
#include <cstdlib>
#include <string>
#include <vector>

#pragma comment(lib, "gdiplus.lib")

namespace
{
void ShowError(const wchar_t* message)
{
    MessageBoxW(NULL, message, L"Spectrogram Context Menu", MB_OK | MB_ICONERROR);
}

bool FindExecutable(const wchar_t* name, std::wstring& path)
{
    DWORD required = SearchPathW(NULL, name, NULL, 0, NULL, NULL);
    if (required == 0)
        return false;

    std::vector<wchar_t> buffer(required + 1);
    DWORD length = SearchPathW(NULL, name, NULL,
        static_cast<DWORD>(buffer.size()), buffer.data(), NULL);
    if (length == 0 || length >= buffer.size())
        return false;

    path.assign(buffer.data(), length);
    return true;
}

bool GetOutputPath(std::wstring& path)
{
    DWORD required = GetTempPathW(0, NULL);
    if (required == 0)
        return false;

    std::vector<wchar_t> buffer(required + 1);
    DWORD length = GetTempPathW(static_cast<DWORD>(buffer.size()), buffer.data());
    if (length == 0 || length >= buffer.size())
        return false;

    path.assign(buffer.data(), length);
    path += L"spec.png";
    return true;
}

int GetEncoderClsid(const wchar_t* format, CLSID* clsid)
{
    using namespace Gdiplus;

    UINT count = 0;
    UINT size = 0;
    GetImageEncodersSize(&count, &size);
    if (size == 0)
        return -1;

    std::vector<BYTE> storage(size);
    ImageCodecInfo* codecs = reinterpret_cast<ImageCodecInfo*>(storage.data());
    if (GetImageEncoders(count, size, codecs) != Ok)
        return -1;

    for (UINT i = 0; i < count; ++i)
    {
        if (wcscmp(codecs[i].MimeType, format) == 0)
        {
            *clsid = codecs[i].Clsid;
            return static_cast<int>(i);
        }
    }
    return -1;
}

void TrimAscii(std::string& value)
{
    size_t first = value.find_first_not_of(" \t\r");
    if (first == std::string::npos)
    {
        value.clear();
        return;
    }
    size_t last = value.find_last_not_of(" \t\r");
    value = value.substr(first, last - first + 1);
}

bool TryParseDouble(const std::string& text, double& value)
{
    char* end = NULL;
    value = std::strtod(text.c_str(), &end);
    if (end == text.c_str())
        return false;
    while (*end == ' ' || *end == '\t')
        ++end;
    return *end == '\0';
}

bool TryParseInt(const std::string& text, int& value)
{
    char* end = NULL;
    long parsed = std::strtol(text.c_str(), &end, 10);
    if (end == text.c_str())
        return false;
    while (*end == ' ' || *end == '\t')
        ++end;
    if (*end != '\0')
        return false;
    value = static_cast<int>(parsed);
    return true;
}

std::wstring ToWideString(const std::string& text)
{
    if (text.empty())
        return L"";

    UINT codePage = CP_UTF8;
    int length = MultiByteToWideChar(codePage, MB_ERR_INVALID_CHARS,
        text.c_str(), -1, NULL, 0);
    if (length <= 0)
    {
        codePage = CP_ACP;
        length = MultiByteToWideChar(codePage, 0, text.c_str(), -1, NULL, 0);
    }
    if (length <= 0)
        return L"";

    std::vector<wchar_t> buffer(length);
    MultiByteToWideChar(codePage, 0, text.c_str(), -1, buffer.data(), length);
    return buffer.data();
}

bool ReadProcessOutput(const std::wstring& command, DWORD timeout,
    std::string& output)
{
    SECURITY_ATTRIBUTES security = {};
    security.nLength = sizeof(security);
    security.bInheritHandle = TRUE;

    HANDLE readPipe = NULL;
    HANDLE writePipe = NULL;
    if (!CreatePipe(&readPipe, &writePipe, &security, 0))
        return false;
    SetHandleInformation(readPipe, HANDLE_FLAG_INHERIT, 0);

    HANDLE nullOutput = CreateFileW(L"NUL", GENERIC_WRITE,
        FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING,
        FILE_ATTRIBUTE_NORMAL, NULL);
    if (nullOutput == INVALID_HANDLE_VALUE)
    {
        CloseHandle(readPipe);
        CloseHandle(writePipe);
        return false;
    }

    STARTUPINFOW startup = { sizeof(startup) };
    startup.dwFlags = STARTF_USESHOWWINDOW | STARTF_USESTDHANDLES;
    startup.wShowWindow = SW_HIDE;
    startup.hStdOutput = writePipe;
    startup.hStdError = nullOutput;
    startup.hStdInput = GetStdHandle(STD_INPUT_HANDLE);

    PROCESS_INFORMATION process = {};
    std::vector<wchar_t> mutableCommand(command.begin(), command.end());
    mutableCommand.push_back(L'\0');
    BOOL started = CreateProcessW(NULL, mutableCommand.data(), NULL, NULL, TRUE,
        CREATE_NO_WINDOW, NULL, NULL, &startup, &process);

    CloseHandle(writePipe);
    CloseHandle(nullOutput);
    if (!started)
    {
        CloseHandle(readPipe);
        return false;
    }

    DWORD waitResult = WaitForSingleObject(process.hProcess, timeout);
    if (waitResult == WAIT_TIMEOUT)
    {
        TerminateProcess(process.hProcess, 1);
        WaitForSingleObject(process.hProcess, 2000);
    }

    char buffer[1024];
    for (;;)
    {
        DWORD bytesRead = 0;
        if (!ReadFile(readPipe, buffer, sizeof(buffer), &bytesRead, NULL) || bytesRead == 0)
            break;
        output.append(buffer, bytesRead);
    }

    DWORD exitCode = 1;
    GetExitCodeProcess(process.hProcess, &exitCode);
    CloseHandle(readPipe);
    CloseHandle(process.hProcess);
    CloseHandle(process.hThread);
    return waitResult == WAIT_OBJECT_0 && exitCode == 0;
}

void GetAudioInfo(const std::wstring& audioFile, double& sampleRate,
    double& duration, std::wstring& codec, int& bitrate, std::wstring& format)
{
    sampleRate = 44100.0;
    duration = 0.0;
    codec.clear();
    bitrate = 0;
    format.clear();

    std::wstring ffprobePath;
    if (!FindExecutable(L"ffprobe.exe", ffprobePath))
        return;

    std::wstring command = L"\"" + ffprobePath +
        L"\" -v error -select_streams a:0"
        L" -show_entries stream=sample_rate,codec_name:format=duration,bit_rate,format_name"
        L" -of default=noprint_wrappers=1 \"" + audioFile + L"\"";

    std::string output;
    if (!ReadProcessOutput(command, 5000, output))
        return;

    size_t position = 0;
    while (position <= output.size())
    {
        size_t newline = output.find('\n', position);
        std::string line = newline == std::string::npos
            ? output.substr(position)
            : output.substr(position, newline - position);
        TrimAscii(line);

        size_t separator = line.find('=');
        if (separator != std::string::npos)
        {
            std::string key = line.substr(0, separator);
            std::string value = line.substr(separator + 1);
            TrimAscii(key);
            TrimAscii(value);

            if (key == "sample_rate")
            {
                double parsed = 0.0;
                if (TryParseDouble(value, parsed) && parsed > 0.0)
                    sampleRate = parsed;
            }
            else if (key == "duration")
            {
                double parsed = 0.0;
                if (TryParseDouble(value, parsed) && parsed >= 0.0)
                    duration = parsed;
            }
            else if (key == "bit_rate")
            {
                int parsed = 0;
                if (TryParseInt(value, parsed) && parsed > 0)
                    bitrate = parsed;
            }
            else if (key == "codec_name")
                codec = ToWideString(value);
            else if (key == "format_name")
                format = ToWideString(value);
        }

        if (newline == std::string::npos)
            break;
        position = newline + 1;
    }
}

void DrawLabels(const std::wstring& pngPath, double sampleRate, double duration,
    const std::wstring& codec, int bitrate, const std::wstring& format,
    const std::wstring& filename)
{
    using namespace Gdiplus;

    GdiplusStartupInput startup;
    ULONG_PTR token = 0;
    if (GdiplusStartup(&token, &startup, NULL) != Ok)
        return;

    std::wstring temporaryPath = pngPath + L".tmp";
    DeleteFileW(temporaryPath.c_str());
    bool saved = false;

    {
        Bitmap source(pngPath.c_str());
        if (source.GetLastStatus() == Ok)
        {
            int sourceWidth = static_cast<int>(source.GetWidth());
            int sourceHeight = static_cast<int>(source.GetHeight());
            const int left = 54;
            const int right = 6;
            const int top = 28;
            const int bottom = 22;
            int totalWidth = left + sourceWidth + right;
            int totalHeight = top + sourceHeight + bottom;

            Bitmap output(totalWidth, totalHeight, PixelFormat32bppARGB);
            Graphics graphics(&output);
            graphics.SetSmoothingMode(SmoothingModeNone);
            graphics.SetTextRenderingHint(TextRenderingHintSingleBitPerPixelGridFit);
            graphics.Clear(Color(255, 0, 0, 0));
            graphics.DrawImage(&source, left, top, sourceWidth, sourceHeight);

            Pen border(Color(255, 100, 100, 100), 1.0f);
            Pen ticks(Color(255, 130, 130, 130), 1.0f);
            SolidBrush textBrush(Color(255, 210, 210, 210));
            graphics.DrawRectangle(&border, left, top, sourceWidth - 1, sourceHeight - 1);

            wchar_t sampleRateText[32];
            if (static_cast<int>(sampleRate) % 1000 == 0)
                swprintf_s(sampleRateText, L"SR %.0fk", sampleRate / 1000.0);
            else
                swprintf_s(sampleRateText, L"SR %.1fk", sampleRate / 1000.0);

            wchar_t bitrateText[32];
            if (bitrate > 0)
                swprintf_s(bitrateText, L"%d kbps", bitrate / 1000);
            else
                wcscpy_s(bitrateText, L"? kbps");

            std::wstring shortFormat = format;
            size_t comma = shortFormat.find(L',');
            if (comma != std::wstring::npos)
                shortFormat.resize(comma);
            if (shortFormat.empty())
                shortFormat = L"?";

            std::wstring info = sampleRateText;
            info += L"  |  ";
            info += bitrateText;
            info += L"  |  " + shortFormat + L"  |  ";
            info += codec.empty() ? L"?" : codec;
            info += L"  |  " + filename;

            Font infoFont(L"Segoe UI", 9, FontStyleRegular);
            RectF infoBounds;
            graphics.MeasureString(info.c_str(), -1, &infoFont, PointF(0, 0), &infoBounds);
            float infoY = (top - infoBounds.Height) * 0.5f;
            graphics.DrawString(info.c_str(), -1, &infoFont,
                PointF(static_cast<float>(left), infoY < 1.0f ? 1.0f : infoY), &textBrush);

            if (sampleRate <= 1000.0)
                sampleRate = 44100.0;
            double maximumFrequency = sampleRate / 2.0;
            double rawStep = maximumFrequency / 18.0;
            double magnitude = std::pow(10.0, std::floor(std::log10(rawStep)));
            double normalized = rawStep / magnitude;
            double frequencyStep = normalized < 1.5 ? magnitude
                : normalized < 3.5 ? 2.0 * magnitude
                : normalized < 7.5 ? 5.0 * magnitude : 10.0 * magnitude;

            Font axisFont(L"Segoe UI", 8, FontStyleRegular);
            int tickCount = static_cast<int>(std::floor((maximumFrequency - 1.0) / frequencyStep));
            for (int i = 1; i <= tickCount; ++i)
            {
                double frequency = frequencyStep * i;
                int y = top + static_cast<int>(std::round(
                    sourceHeight * (1.0 - frequency / maximumFrequency)));
                graphics.DrawLine(&ticks, left - 4, y, left, y);

                wchar_t label[32];
                if (frequency >= 1000.0)
                    swprintf_s(label, L"%.0fk", frequency / 1000.0);
                else
                    swprintf_s(label, L"%.0f", frequency);
                RectF bounds;
                graphics.MeasureString(label, -1, &axisFont, PointF(0, 0), &bounds);
                graphics.DrawString(label, -1, &axisFont,
                    PointF(24.0f, static_cast<float>(y) - bounds.Height * 0.5f), &textBrush);
            }

            if (duration > 0.0)
            {
                const int divisions = 16;
                for (int i = 0; i <= divisions; ++i)
                {
                    double time = duration * i / divisions;
                    int x = left + static_cast<int>(std::round(
                        static_cast<double>(sourceWidth) * i / divisions));
                    graphics.DrawLine(&ticks, x, top + sourceHeight, x, top + sourceHeight + 4);

                    wchar_t label[32];
                    swprintf_s(label, L"%d:%02d", static_cast<int>(time / 60.0),
                        static_cast<int>(time) % 60);
                    RectF bounds;
                    graphics.MeasureString(label, -1, &axisFont, PointF(0, 0), &bounds);
                    float labelX = static_cast<float>(x) - bounds.Width * 0.5f;
                    if (labelX < 0.0f)
                        labelX = 0.0f;
                    if (labelX + bounds.Width > totalWidth)
                        labelX = static_cast<float>(totalWidth) - bounds.Width;
                    graphics.DrawString(label, -1, &axisFont,
                        PointF(labelX, static_cast<float>(top + sourceHeight + 6)), &textBrush);
                }
            }

            CLSID encoder;
            if (GetEncoderClsid(L"image/png", &encoder) >= 0)
                saved = output.Save(temporaryPath.c_str(), &encoder) == Ok;
        }
    }

    GdiplusShutdown(token);
    if (saved)
    {
        if (!MoveFileExW(temporaryPath.c_str(), pngPath.c_str(),
            MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH))
            DeleteFileW(temporaryPath.c_str());
    }
    else
        DeleteFileW(temporaryPath.c_str());
}

bool RunFFmpeg(const std::wstring& command)
{
    STARTUPINFOW startup = { sizeof(startup) };
    startup.dwFlags = STARTF_USESHOWWINDOW;
    startup.wShowWindow = SW_HIDE;

    PROCESS_INFORMATION process = {};
    std::vector<wchar_t> mutableCommand(command.begin(), command.end());
    mutableCommand.push_back(L'\0');
    if (!CreateProcessW(NULL, mutableCommand.data(), NULL, NULL, FALSE,
        CREATE_NO_WINDOW, NULL, NULL, &startup, &process))
        return false;

    DWORD waitResult = WaitForSingleObject(process.hProcess, FFMPEG_TIMEOUT_MS);
    if (waitResult == WAIT_TIMEOUT)
    {
        TerminateProcess(process.hProcess, 1);
        WaitForSingleObject(process.hProcess, 2000);
    }

    DWORD exitCode = 1;
    GetExitCodeProcess(process.hProcess, &exitCode);
    CloseHandle(process.hProcess);
    CloseHandle(process.hThread);
    return waitResult == WAIT_OBJECT_0 && exitCode == 0;
}
}

bool GenerateAndShowSpectrogram(const std::wstring& audioFile)
{
    std::wstring ffmpegPath;
    if (!FindExecutable(L"ffmpeg.exe", ffmpegPath))
    {
        ShowError(L"ffmpeg.exe was not found. Make sure FFmpeg is available in PATH.");
        return false;
    }

    std::wstring outputPath;
    if (!GetOutputPath(outputPath))
    {
        ShowError(L"The temporary directory could not be located.");
        return false;
    }
    DeleteFileW(outputPath.c_str());

    std::wstring command = L"\"" + ffmpegPath + L"\" -nostdin -y -i \"" +
        audioFile + L"\" -lavfi \"showspectrumpic=s=" +
        std::to_wstring(SPECTROGRAM_WIDTH) + L"x" +
        std::to_wstring(SPECTROGRAM_HEIGHT) + L":mode=" + SPECTROGRAM_MODE +
        L":color=" + SPECTROGRAM_COLOR + L":scale=" + SPECTROGRAM_SCALE +
        L":legend=0,unsharp=" + SPECTROGRAM_UNSHARP + L"\" \"" + outputPath + L"\"";

    if (!RunFFmpeg(command) || !PathFileExistsW(outputPath.c_str()))
    {
        ShowError(L"The spectrogram could not be created. Check the audio file format.");
        return false;
    }

    double sampleRate = 44100.0;
    double duration = 0.0;
    std::wstring codec;
    std::wstring format;
    int bitrate = 0;
    GetAudioInfo(audioFile, sampleRate, duration, codec, bitrate, format);

    const wchar_t* filenamePart = PathFindFileNameW(audioFile.c_str());
    DrawLabels(outputPath, sampleRate, duration, codec, bitrate, format,
        filenamePart != NULL ? filenamePart : audioFile);

    HINSTANCE openResult = ShellExecuteW(NULL, L"open", outputPath.c_str(),
        NULL, NULL, SW_SHOW);
    if (reinterpret_cast<INT_PTR>(openResult) <= 32)
    {
        ShowError(L"The generated image could not be opened.");
        return false;
    }
    return true;
}
