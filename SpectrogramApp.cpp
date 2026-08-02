#include "SpectrogramContextMenu.h"
#include "resource.h"

#include <windows.h>
#include <shellapi.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <string>
#include <vector>

namespace
{
const wchar_t* const kAppName = L"Spectrogram Context Menu";
const wchar_t* const kExeName = L"SpectrogramContextMenu.exe";
const wchar_t* const kInstallSubdirectory =
    L"Programs\\SpectrogramContextMenu";
const wchar_t* const kVerbName = L"SpectrogramContextMenu";
const wchar_t* const kMenuText = L"\u041f\u043e\u043a\u0430\u0437\u0430\u0442\u044c \u0441\u043f\u0435\u043a\u0442\u0440\u043e\u0433\u0440\u0430\u043c\u043c\u0443";

const wchar_t* const kAudioExtensions[] = {
    L".mp3", L".flac", L".wav", L".ogg", L".m4a",
    L".aac", L".wma", L".opus", L".ape", L".ac3"
};

void ShowError(const std::wstring& message)
{
    MessageBoxW(NULL, message.c_str(), kAppName, MB_OK | MB_ICONERROR);
}

void ShowInfo(const std::wstring& message)
{
    MessageBoxW(NULL, message.c_str(), kAppName, MB_OK | MB_ICONINFORMATION);
}

std::wstring GetLastErrorMessage(const wchar_t* operation)
{
    DWORD error = GetLastError();
    wchar_t* systemMessage = NULL;
    FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM |
        FORMAT_MESSAGE_IGNORE_INSERTS, NULL, error, 0,
        reinterpret_cast<wchar_t*>(&systemMessage), 0, NULL);

    std::wstring result(operation);
    result += L"\n\n";
    if (systemMessage != NULL)
    {
        result += systemMessage;
        LocalFree(systemMessage);
    }
    else
    {
        result += L"Windows error ";
        result += std::to_wstring(error);
    }
    return result;
}

bool GetModulePath(std::wstring& path)
{
    std::vector<wchar_t> buffer(512);
    for (;;)
    {
        DWORD length = GetModuleFileNameW(NULL, buffer.data(),
            static_cast<DWORD>(buffer.size()));
        if (length == 0)
            return false;
        if (length < buffer.size() - 1)
        {
            path.assign(buffer.data(), length);
            return true;
        }
        buffer.resize(buffer.size() * 2);
    }
}

bool FindInPath(const wchar_t* filename, std::wstring& path)
{
    DWORD required = SearchPathW(NULL, filename, NULL, 0, NULL, NULL);
    if (required == 0)
        return false;

    std::vector<wchar_t> buffer(required + 1);
    DWORD length = SearchPathW(NULL, filename, NULL,
        static_cast<DWORD>(buffer.size()), buffer.data(), NULL);
    if (length == 0 || length >= buffer.size())
        return false;

    path.assign(buffer.data(), length);
    return true;
}

bool CheckFFmpeg(std::wstring& missingFiles)
{
    std::wstring path;
    bool hasFFmpeg = FindInPath(L"ffmpeg.exe", path);
    bool hasFFprobe = FindInPath(L"ffprobe.exe", path);
    if (hasFFmpeg && hasFFprobe)
        return true;

    if (!hasFFmpeg)
        missingFiles = L"ffmpeg.exe";
    if (!hasFFprobe)
    {
        if (!missingFiles.empty())
            missingFiles += L", ";
        missingFiles += L"ffprobe.exe";
    }
    return false;
}

bool GetInstallPaths(std::wstring& directory, std::wstring& executable)
{
    wchar_t localAppData[MAX_PATH];
    if (FAILED(SHGetFolderPathW(NULL, CSIDL_LOCAL_APPDATA | CSIDL_FLAG_CREATE,
        NULL, SHGFP_TYPE_CURRENT, localAppData)))
        return false;

    directory = localAppData;
    directory += L"\\";
    directory += kInstallSubdirectory;
    executable = directory + L"\\" + kExeName;
    return true;
}

bool SetStringValue(HKEY root, const std::wstring& subkey,
    const wchar_t* valueName, const std::wstring& value)
{
    HKEY key = NULL;
    LONG status = RegCreateKeyExW(root, subkey.c_str(), 0, NULL,
        REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, NULL, &key, NULL);
    if (status != ERROR_SUCCESS)
    {
        SetLastError(status);
        return false;
    }

    status = RegSetValueExW(key, valueName, 0, REG_SZ,
        reinterpret_cast<const BYTE*>(value.c_str()),
        static_cast<DWORD>((value.size() + 1) * sizeof(wchar_t)));
    RegCloseKey(key);
    if (status != ERROR_SUCCESS)
    {
        SetLastError(status);
        return false;
    }
    return true;
}

bool SetDwordValue(HKEY root, const std::wstring& subkey,
    const wchar_t* valueName, DWORD value)
{
    HKEY key = NULL;
    LONG status = RegCreateKeyExW(root, subkey.c_str(), 0, NULL,
        REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, NULL, &key, NULL);
    if (status == ERROR_SUCCESS)
    {
        status = RegSetValueExW(key, valueName, 0, REG_DWORD,
            reinterpret_cast<const BYTE*>(&value), sizeof(value));
        RegCloseKey(key);
    }
    if (status != ERROR_SUCCESS)
    {
        SetLastError(status);
        return false;
    }
    return true;
}

void RemoveRegistryEntries()
{
    for (size_t i = 0; i < ARRAYSIZE(kAudioExtensions); ++i)
    {
        std::wstring key = L"Software\\Classes\\SystemFileAssociations\\";
        key += kAudioExtensions[i];
        key += L"\\shell\\";
        key += kVerbName;
        RegDeleteTreeW(HKEY_CURRENT_USER, key.c_str());
    }
    RegDeleteTreeW(HKEY_CURRENT_USER,
        L"Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\SpectrogramContextMenu");
}

bool RegisterApplication(const std::wstring& installedExe)
{
    RemoveRegistryEntries();

    std::wstring quotedExe = L"\"" + installedExe + L"\"";
    std::wstring command = quotedExe + L" \"%1\"";
    std::wstring icon = quotedExe + L",0";

    for (size_t i = 0; i < ARRAYSIZE(kAudioExtensions); ++i)
    {
        std::wstring verbKey = L"Software\\Classes\\SystemFileAssociations\\";
        verbKey += kAudioExtensions[i];
        verbKey += L"\\shell\\";
        verbKey += kVerbName;

        if (!SetStringValue(HKEY_CURRENT_USER, verbKey, NULL, kMenuText) ||
            !SetStringValue(HKEY_CURRENT_USER, verbKey, L"MUIVerb", kMenuText) ||
            !SetStringValue(HKEY_CURRENT_USER, verbKey, L"Icon", icon) ||
            !SetStringValue(HKEY_CURRENT_USER, verbKey + L"\\command", NULL, command))
        {
            RemoveRegistryEntries();
            return false;
        }
    }

    const std::wstring uninstallKey =
        L"Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\SpectrogramContextMenu";
    if (!SetStringValue(HKEY_CURRENT_USER, uninstallKey, L"DisplayName", kAppName) ||
        !SetStringValue(HKEY_CURRENT_USER, uninstallKey, L"DisplayIcon", icon) ||
        !SetStringValue(HKEY_CURRENT_USER, uninstallKey, L"DisplayVersion", L"4.0.0") ||
        !SetStringValue(HKEY_CURRENT_USER, uninstallKey, L"Publisher", L"Spectrogram Context Menu") ||
        !SetStringValue(HKEY_CURRENT_USER, uninstallKey, L"InstallLocation",
            installedExe.substr(0, installedExe.find_last_of(L"\\/"))) ||
        !SetStringValue(HKEY_CURRENT_USER, uninstallKey, L"UninstallString",
            quotedExe + L" /uninstall") ||
        !SetStringValue(HKEY_CURRENT_USER, uninstallKey, L"QuietUninstallString",
            quotedExe + L" /uninstall /silent") ||
        !SetDwordValue(HKEY_CURRENT_USER, uninstallKey, L"NoModify", 1) ||
        !SetDwordValue(HKEY_CURRENT_USER, uninstallKey, L"NoRepair", 0))
    {
        RemoveRegistryEntries();
        return false;
    }
    return true;
}

void NotifyExplorer()
{
    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);
}

bool InstallOrRepair(bool repair)
{
    std::wstring missing;
    if (!CheckFFmpeg(missing))
    {
        ShowError(L"\u0423\u0441\u0442\u0430\u043d\u043e\u0432\u043a\u0430 \u043d\u0435 \u0432\u044b\u043f\u043e\u043b\u043d\u0435\u043d\u0430. \u0412 PATH \u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d\u044b:\n\n" + missing +
            L"\n\n\u0414\u043e\u0431\u0430\u0432\u044c\u0442\u0435 FFmpeg \u0432 PATH \u0438 \u0437\u0430\u043f\u0443\u0441\u0442\u0438\u0442\u0435 \u044d\u0442\u043e\u0442 EXE \u0441\u043d\u043e\u0432\u0430.");
        return false;
    }

    std::wstring sourceExe;
    std::wstring installDirectory;
    std::wstring installedExe;
    if (!GetModulePath(sourceExe) || !GetInstallPaths(installDirectory, installedExe))
    {
        ShowError(L"\u041d\u0435 \u0443\u0434\u0430\u043b\u043e\u0441\u044c \u043e\u043f\u0440\u0435\u0434\u0435\u043b\u0438\u0442\u044c \u043f\u0430\u043f\u043a\u0443 \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u043a\u0438.");
        return false;
    }

    if (SHCreateDirectoryExW(NULL, installDirectory.c_str(), NULL) != ERROR_SUCCESS &&
        GetFileAttributesW(installDirectory.c_str()) == INVALID_FILE_ATTRIBUTES)
    {
        ShowError(GetLastErrorMessage(L"\u041d\u0435 \u0443\u0434\u0430\u043b\u043e\u0441\u044c \u0441\u043e\u0437\u0434\u0430\u0442\u044c \u043f\u0430\u043f\u043a\u0443 \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u043a\u0438."));
        return false;
    }

    if (_wcsicmp(sourceExe.c_str(), installedExe.c_str()) != 0 &&
        !CopyFileW(sourceExe.c_str(), installedExe.c_str(), FALSE))
    {
        ShowError(GetLastErrorMessage(L"\u041d\u0435 \u0443\u0434\u0430\u043b\u043e\u0441\u044c \u0441\u043a\u043e\u043f\u0438\u0440\u043e\u0432\u0430\u0442\u044c \u043f\u0440\u043e\u0433\u0440\u0430\u043c\u043c\u0443."));
        return false;
    }

    if (!RegisterApplication(installedExe))
    {
        ShowError(GetLastErrorMessage(L"\u041d\u0435 \u0443\u0434\u0430\u043b\u043e\u0441\u044c \u0437\u0430\u0440\u0435\u0433\u0438\u0441\u0442\u0440\u0438\u0440\u043e\u0432\u0430\u0442\u044c \u043a\u043e\u043d\u0442\u0435\u043a\u0441\u0442\u043d\u043e\u0435 \u043c\u0435\u043d\u044e."));
        return false;
    }

    NotifyExplorer();
    ShowInfo(repair
        ? L"\u0423\u0441\u0442\u0430\u043d\u043e\u0432\u043a\u0430 \u0432\u043e\u0441\u0441\u0442\u0430\u043d\u043e\u0432\u043b\u0435\u043d\u0430."
        : L"\u041f\u0440\u043e\u0433\u0440\u0430\u043c\u043c\u0430 \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u043b\u0435\u043d\u0430.\n\n\u041f\u0443\u043d\u043a \u00ab\u041f\u043e\u043a\u0430\u0437\u0430\u0442\u044c \u0441\u043f\u0435\u043a\u0442\u0440\u043e\u0433\u0440\u0430\u043c\u043c\u0443\u00bb \u0434\u043e\u0431\u0430\u0432\u043b\u0435\u043d \u0432 \u043a\u043e\u043d\u0442\u0435\u043a\u0441\u0442\u043d\u043e\u0435 \u043c\u0435\u043d\u044e \u0430\u0443\u0434\u0438\u043e\u0444\u0430\u0439\u043b\u043e\u0432.");
    return true;
}

bool WriteSelfDeleteScript(const std::wstring& executable,
    const std::wstring& directory, std::wstring& scriptPath)
{
    wchar_t tempPath[MAX_PATH];
    if (GetTempPathW(ARRAYSIZE(tempPath), tempPath) == 0)
        return false;

    scriptPath = tempPath;
    scriptPath += L"SpectrogramContextMenu-uninstall.cmd";
    HANDLE file = CreateFileW(scriptPath.c_str(), GENERIC_WRITE, 0, NULL,
        CREATE_ALWAYS, FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_TEMPORARY, NULL);
    if (file == INVALID_HANDLE_VALUE)
        return false;

    std::wstring script =
        L"@chcp 65001 >nul\r\n"
        L"@echo off\r\n"
        L":wait\r\n"
        L"del /f /q \"" + executable + L"\" >nul 2>&1\r\n"
        L"if exist \"" + executable + L"\" (\r\n"
        L"  timeout /t 1 /nobreak >nul\r\n"
        L"  goto wait\r\n"
        L")\r\n"
        L"rmdir \"" + directory + L"\" >nul 2>&1\r\n"
        L"del /f /q \"%~f0\"\r\n";

    int utf8Length = WideCharToMultiByte(CP_UTF8, 0, script.c_str(),
        static_cast<int>(script.size()), NULL, 0, NULL, NULL);
    std::vector<char> utf8(utf8Length);
    WideCharToMultiByte(CP_UTF8, 0, script.c_str(),
        static_cast<int>(script.size()), utf8.data(), utf8Length, NULL, NULL);

    DWORD written = 0;
    BOOL ok = WriteFile(file, utf8.data(), static_cast<DWORD>(utf8.size()),
        &written, NULL);
    CloseHandle(file);
    return ok && written == utf8.size();
}

bool Uninstall(bool silent)
{
    RemoveRegistryEntries();
    NotifyExplorer();

    std::wstring currentExe;
    std::wstring installDirectory;
    std::wstring installedExe;
    if (GetModulePath(currentExe) && GetInstallPaths(installDirectory, installedExe))
    {
        if (_wcsicmp(currentExe.c_str(), installedExe.c_str()) == 0)
        {
            std::wstring scriptPath;
            if (!WriteSelfDeleteScript(installedExe, installDirectory, scriptPath))
            {
                ShowError(GetLastErrorMessage(L"\u041d\u0435 \u0443\u0434\u0430\u043b\u043e\u0441\u044c \u0437\u0430\u043f\u0443\u0441\u0442\u0438\u0442\u044c \u0441\u0430\u043c\u043e\u0443\u0434\u0430\u043b\u0435\u043d\u0438\u0435."));
                return false;
            }
            HINSTANCE launchResult = ShellExecuteW(NULL, L"open", scriptPath.c_str(),
                NULL, NULL, SW_HIDE);
            if (reinterpret_cast<INT_PTR>(launchResult) <= 32)
            {
                ShowError(L"\u041d\u0435 \u0443\u0434\u0430\u043b\u043e\u0441\u044c \u0437\u0430\u043f\u0443\u0441\u0442\u0438\u0442\u044c \u0441\u0430\u043c\u043e\u0443\u0434\u0430\u043b\u0435\u043d\u0438\u0435.");
                return false;
            }
        }
        else
        {
            DeleteFileW(installedExe.c_str());
            RemoveDirectoryW(installDirectory.c_str());
        }
    }

    if (!silent)
        ShowInfo(L"\u041f\u0440\u043e\u0433\u0440\u0430\u043c\u043c\u0430 \u0443\u0434\u0430\u043b\u0435\u043d\u0430 \u0434\u043b\u044f \u0442\u0435\u043a\u0443\u0449\u0435\u0433\u043e \u043f\u043e\u043b\u044c\u0437\u043e\u0432\u0430\u0442\u0435\u043b\u044f.");
    return true;
}

bool IsSupportedAudioFile(const std::wstring& path)
{
    const wchar_t* extension = PathFindExtensionW(path.c_str());
    for (size_t i = 0; i < ARRAYSIZE(kAudioExtensions); ++i)
    {
        if (_wcsicmp(extension, kAudioExtensions[i]) == 0)
            return true;
    }
    return false;
}

bool IsOption(const wchar_t* value, const wchar_t* option)
{
    return _wcsicmp(value, option) == 0;
}
}

int APIENTRY wWinMain(HINSTANCE, HINSTANCE, LPWSTR, int)
{
    int argc = 0;
    LPWSTR* argv = CommandLineToArgvW(GetCommandLineW(), &argc);
    if (argv == NULL)
    {
        ShowError(L"\u041d\u0435 \u0443\u0434\u0430\u043b\u043e\u0441\u044c \u0440\u0430\u0437\u043e\u0431\u0440\u0430\u0442\u044c \u043a\u043e\u043c\u0430\u043d\u0434\u043d\u0443\u044e \u0441\u0442\u0440\u043e\u043a\u0443.");
        return 1;
    }

    bool result = false;
    if (argc == 1)
    {
        result = InstallOrRepair(false);
    }
    else if (IsOption(argv[1], L"/uninstall"))
    {
        bool silent = argc > 2 && IsOption(argv[2], L"/silent");
        result = Uninstall(silent);
    }
    else if (IsOption(argv[1], L"/repair"))
    {
        result = InstallOrRepair(true);
    }
    else
    {
        std::wstring path = argv[1];
        DWORD attributes = GetFileAttributesW(path.c_str());
        if (attributes == INVALID_FILE_ATTRIBUTES ||
            (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
        {
            ShowError(L"\u0410\u0443\u0434\u0438\u043e\u0444\u0430\u0439\u043b \u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d.");
        }
        else if (!IsSupportedAudioFile(path))
        {
            ShowError(L"\u042d\u0442\u043e\u0442 \u0444\u043e\u0440\u043c\u0430\u0442 \u0444\u0430\u0439\u043b\u0430 \u043d\u0435 \u043f\u043e\u0434\u0434\u0435\u0440\u0436\u0438\u0432\u0430\u0435\u0442\u0441\u044f.");
        }
        else
        {
            std::wstring missing;
            if (!CheckFFmpeg(missing))
            {
                ShowError(L"\u0412 PATH \u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d\u044b: " + missing +
                    L".\n\n\u041f\u0440\u043e\u0432\u0435\u0440\u044c\u0442\u0435 \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u043a\u0443 FFmpeg.");
            }
            else
            {
                result = GenerateAndShowSpectrogram(path);
            }
        }
    }

    LocalFree(argv);
    return result ? 0 : 1;
}
