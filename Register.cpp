#include "SpectrogramContextMenu.h"
#include <strsafe.h>

extern HINSTANCE g_hInst;

// Вспомогательная функция для записи в реестр
HRESULT SetRegistryKeyAndValue(LPCWSTR pszSubKey, LPCWSTR pszValueName, LPCWSTR pszData)
{
    HKEY hKey = NULL;
    LONG lResult;

    // Пытаемся создать ключ с полными правами
    lResult = RegCreateKeyExW(HKEY_CLASSES_ROOT, pszSubKey, 0,
        NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE | KEY_WOW64_64KEY, NULL, &hKey, NULL);

    if (lResult != ERROR_SUCCESS)
    {
        // Выводим MessageBox для отладки
        WCHAR szError[512];
        StringCchPrintfW(szError, ARRAYSIZE(szError), 
            L"Failed to create registry key:\n%s\n\nError code: %d", 
            pszSubKey, lResult);
        MessageBoxW(NULL, szError, L"Registry Error", MB_OK | MB_ICONERROR);
        return HRESULT_FROM_WIN32(lResult);
    }

    if (pszData != NULL)
    {
        DWORD cbData = static_cast<DWORD>((wcslen(pszData) + 1) * sizeof(WCHAR));
        lResult = RegSetValueExW(hKey, pszValueName, 0, 
            REG_SZ, reinterpret_cast<const BYTE *>(pszData), cbData);
        
        if (lResult != ERROR_SUCCESS)
        {
            WCHAR szError[512];
            StringCchPrintfW(szError, ARRAYSIZE(szError), 
                L"Failed to set registry value:\n%s\\%s\n\nError code: %d", 
                pszSubKey, pszValueName ? pszValueName : L"(Default)", lResult);
            MessageBoxW(NULL, szError, L"Registry Error", MB_OK | MB_ICONERROR);
        }
    }

    RegCloseKey(hKey);
    return HRESULT_FROM_WIN32(lResult);
}

// Регистрация COM сервера
HRESULT RegisterServer()
{
    HRESULT hr;
    WCHAR szModule[MAX_PATH];
    WCHAR szCLSID[MAX_PATH];

    // Получаем путь к нашей DLL
    if (GetModuleFileNameW(g_hInst, szModule, ARRAYSIZE(szModule)) == 0)
    {
        MessageBoxW(NULL, L"Failed to get module filename", L"Error", MB_OK | MB_ICONERROR);
        return HRESULT_FROM_WIN32(GetLastError());
    }

    // Показываем путь к DLL для отладки
    WCHAR szDebug[512];
    StringCchPrintfW(szDebug, ARRAYSIZE(szDebug), L"Registering DLL:\n%s", szModule);
    MessageBoxW(NULL, szDebug, L"Debug Info", MB_OK | MB_ICONINFORMATION);

    // Конвертируем CLSID в строку
    StringFromGUID2(CLSID_SpectrogramContextMenu, szCLSID, ARRAYSIZE(szCLSID));

    // Регистрируем CLSID
    WCHAR szSubkey[MAX_PATH];
    hr = StringCchPrintfW(szSubkey, ARRAYSIZE(szSubkey), L"CLSID\\%s", szCLSID);
    if (SUCCEEDED(hr))
    {
        hr = SetRegistryKeyAndValue(szSubkey, NULL, L"Spectrogram Context Menu");
        if (SUCCEEDED(hr))
        {
            StringCchPrintfW(szSubkey, ARRAYSIZE(szSubkey), L"CLSID\\%s\\InprocServer32", szCLSID);
            hr = SetRegistryKeyAndValue(szSubkey, NULL, szModule);
            if (SUCCEEDED(hr))
            {
                hr = SetRegistryKeyAndValue(szSubkey, L"ThreadingModel", L"Apartment");
            }
        }
    }

    if (FAILED(hr))
    {
        MessageBoxW(NULL, L"Failed to register CLSID", L"Error", MB_OK | MB_ICONERROR);
        return hr;
    }

    // Регистрируем для всех аудиоформатов
    const wchar_t* extensions[] = {
        L".mp3", L".flac", L".wav", L".ogg", L".m4a", 
        L".aac", L".wma", L".opus", L".ape", L".ac3"
    };

    for (int i = 0; i < ARRAYSIZE(extensions); i++)
    {
        // Регистрируем для расширения
        StringCchPrintfW(szSubkey, ARRAYSIZE(szSubkey), 
                        L"%s\\shellex\\ContextMenuHandlers\\SpectrogramMenu", extensions[i]);
        hr = SetRegistryKeyAndValue(szSubkey, NULL, szCLSID);
        
        if (FAILED(hr))
        {
            StringCchPrintfW(szDebug, ARRAYSIZE(szDebug), 
                L"Failed to register extension: %s", extensions[i]);
            MessageBoxW(NULL, szDebug, L"Warning", MB_OK | MB_ICONWARNING);
        }
        
        // Также регистрируем для ProgID (если есть)
        HKEY hKey;
        WCHAR szProgID[256] = {0};
        DWORD dwSize = sizeof(szProgID);
        
        if (RegOpenKeyExW(HKEY_CLASSES_ROOT, extensions[i], 0, KEY_READ, &hKey) == ERROR_SUCCESS)
        {
            if (RegQueryValueExW(hKey, NULL, NULL, NULL, (LPBYTE)szProgID, &dwSize) == ERROR_SUCCESS && szProgID[0])
            {
                // Нашли ProgID, регистрируем для него
                StringCchPrintfW(szSubkey, ARRAYSIZE(szSubkey), 
                                L"%s\\shellex\\ContextMenuHandlers\\SpectrogramMenu", szProgID);
                SetRegistryKeyAndValue(szSubkey, NULL, szCLSID);
            }
            RegCloseKey(hKey);
        }
    }

    MessageBoxW(NULL, L"Registration completed!", L"Success", MB_OK | MB_ICONINFORMATION);
    return S_OK;
}

// Удаление регистрации COM сервера
HRESULT UnregisterServer()
{
    HRESULT hr = S_OK;
    WCHAR szCLSID[MAX_PATH];
    WCHAR szSubkey[MAX_PATH];

    StringFromGUID2(CLSID_SpectrogramContextMenu, szCLSID, ARRAYSIZE(szCLSID));

    // Удаляем регистрацию CLSID
    StringCchPrintfW(szSubkey, ARRAYSIZE(szSubkey), L"CLSID\\%s", szCLSID);
    RegDeleteTreeW(HKEY_CLASSES_ROOT, szSubkey);

    // Удаляем регистрацию для расширений
    const wchar_t* extensions[] = {
        L".mp3", L".flac", L".wav", L".ogg", L".m4a", 
        L".aac", L".wma", L".opus", L".ape", L".ac3"
    };

    for (int i = 0; i < ARRAYSIZE(extensions); i++)
    {
        StringCchPrintfW(szSubkey, ARRAYSIZE(szSubkey), 
                        L"%s\\shellex\\ContextMenuHandlers\\SpectrogramMenu", extensions[i]);
        RegDeleteTreeW(HKEY_CLASSES_ROOT, szSubkey);
    }

    return hr;
}

// DLL экспорты для регистрации
extern "C" STDAPI DllRegisterServer()
{
    return RegisterServer();
}

extern "C" STDAPI DllUnregisterServer()
{
    return UnregisterServer();
}
