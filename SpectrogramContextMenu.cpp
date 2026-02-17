#include "SpectrogramContextMenu.h"
#include "resource.h"
#include <strsafe.h>
#include <shellapi.h>
#include <commctrl.h>

static LONG g_cDllRef = 0;
HINSTANCE g_hInst = NULL;  // Убрали static чтобы был доступен из других файлов

// Поддерживаемые аудиоформаты
static const wchar_t* g_audioExtensions[] = {
    L".mp3", L".flac", L".wav", L".ogg", L".m4a", 
    L".aac", L".wma", L".opus", L".ape", L".ac3"
};

SpectrogramContextMenu::SpectrogramContextMenu() : m_cRef(1)
{
    InterlockedIncrement(&g_cDllRef);
}

SpectrogramContextMenu::~SpectrogramContextMenu()
{
    InterlockedDecrement(&g_cDllRef);
}

// IUnknown методы
STDMETHODIMP SpectrogramContextMenu::QueryInterface(REFIID riid, void **ppvObject)
{
    if (riid == IID_IUnknown || riid == IID_IShellExtInit)
    {
        *ppvObject = static_cast<IShellExtInit*>(this);
    }
    else if (riid == IID_IContextMenu)
    {
        *ppvObject = static_cast<IContextMenu*>(this);
    }
    else
    {
        *ppvObject = NULL;
        return E_NOINTERFACE;
    }

    AddRef();
    return S_OK;
}

STDMETHODIMP_(ULONG) SpectrogramContextMenu::AddRef()
{
    return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) SpectrogramContextMenu::Release()
{
    ULONG cRef = InterlockedDecrement(&m_cRef);
    if (cRef == 0)
    {
        delete this;
    }
    return cRef;
}

// IShellExtInit методы
STDMETHODIMP SpectrogramContextMenu::Initialize(LPCITEMIDLIST pidlFolder, LPDATAOBJECT pDataObj, HKEY hKeyProgID)
{
    if (pDataObj == NULL)
    {
        return E_INVALIDARG;
    }

    FORMATETC fmt = { CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
    STGMEDIUM stg = { TYMED_HGLOBAL };

    if (FAILED(pDataObj->GetData(&fmt, &stg)))
    {
        return E_INVALIDARG;
    }

    HDROP hDrop = static_cast<HDROP>(GlobalLock(stg.hGlobal));
    if (hDrop == NULL)
    {
        ReleaseStgMedium(&stg);
        return E_INVALIDARG;
    }

    UINT nFiles = DragQueryFileW(hDrop, 0xFFFFFFFF, NULL, 0);
    m_selectedFiles.clear();

    for (UINT i = 0; i < nFiles; i++)
    {
        WCHAR szFile[MAX_PATH];
        if (DragQueryFileW(hDrop, i, szFile, ARRAYSIZE(szFile)))
        {
            if (IsAudioFile(szFile))
            {
                m_selectedFiles.push_back(szFile);
            }
        }
    }

    GlobalUnlock(stg.hGlobal);
    ReleaseStgMedium(&stg);

    return m_selectedFiles.empty() ? E_FAIL : S_OK;
}

bool SpectrogramContextMenu::IsAudioFile(const std::wstring& filename)
{
    std::wstring ext = PathFindExtensionW(filename.c_str());
    
    // Преобразуем в нижний регистр
    for (wchar_t& c : ext)
    {
        c = towlower(c);
    }

    for (int i = 0; i < ARRAYSIZE(g_audioExtensions); i++)
    {
        if (ext == g_audioExtensions[i])
        {
            return true;
        }
    }
    return false;
}

// IContextMenu методы
STDMETHODIMP SpectrogramContextMenu::QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags)
{
    if (uFlags & CMF_DEFAULTONLY)
    {
        return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, 0);
    }

    if (m_selectedFiles.empty())
    {
        return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, 0);
    }

    // Пытаемся загрузить нашу собственную иконку
    HICON hIcon = (HICON)LoadImageW(g_hInst, MAKEINTRESOURCE(IDI_SPECTROGRAM), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR);
    if (!hIcon) {
        // Если не удалось загрузить нашу иконку, пробуем системную
        hIcon = (HICON)LoadImageW(NULL, IDI_APPLICATION, IMAGE_ICON, 16, 16, LR_SHARED);
        if (!hIcon) {
            // Если не удалось загрузить даже системную иконку, используем стандартный вид
            MENUITEMINFOW mii = { sizeof(mii) };
            mii.fMask = MIIM_STRING | MIIM_ID | MIIM_STATE;
            mii.wID = idCmdFirst;
            mii.fState = MFS_ENABLED;
            mii.dwTypeData = const_cast<LPWSTR>(L"SHOW SPECTROGRAM");

            if (InsertMenuItemW(hmenu, indexMenu, TRUE, &mii))
            {
                return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, 1);
            }

            return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, 0);
        }
    }

    // Создаем bitmap из иконки
    HDC hdcScreen = GetDC(NULL);
    HDC hdcMem = CreateCompatibleDC(hdcScreen);
    HBITMAP hbm = CreateCompatibleBitmap(hdcScreen, 16, 16);
    HGDIOBJ hOldBitmap = SelectObject(hdcMem, hbm);

    // Рисуем иконку на bitmap
    DrawIconEx(hdcMem, 0, 0, hIcon, 16, 16, 0, NULL, DI_NORMAL);

    // Восстанавливаем контекст устройства
    SelectObject(hdcMem, hOldBitmap);
    DeleteDC(hdcMem);
    ReleaseDC(NULL, hdcScreen);

    MENUITEMINFOW mii = { sizeof(mii) };
    mii.fMask = MIIM_STRING | MIIM_ID | MIIM_STATE | MIIM_BITMAP;
    mii.wID = idCmdFirst;
    mii.fState = MFS_ENABLED;
    mii.dwTypeData = const_cast<LPWSTR>(L"SHOW SPECTROGRAM");
    mii.hbmpItem = hbm;

    if (InsertMenuItemW(hmenu, indexMenu, TRUE, &mii))
    {
        return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, 1);
    }

    // Если не удалось добавить элемент с иконкой, пробуем без иконки
    DeleteObject(hbm);
    mii.fMask = MIIM_STRING | MIIM_ID | MIIM_STATE;
    mii.hbmpItem = NULL;

    if (InsertMenuItemW(hmenu, indexMenu, TRUE, &mii))
    {
        return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, 1);
    }

    return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, 0);
}

STDMETHODIMP SpectrogramContextMenu::InvokeCommand(LPCMINVOKECOMMANDINFO pici)
{
    if (HIWORD(pici->lpVerb))
    {
        return E_INVALIDARG;
    }

    if (LOWORD(pici->lpVerb) != 0)
    {
        return E_INVALIDARG;
    }

    if (!m_selectedFiles.empty())
    {
        GenerateAndShowSpectrogram(m_selectedFiles[0]);
    }

    return S_OK;
}

STDMETHODIMP SpectrogramContextMenu::GetCommandString(UINT_PTR idCmd, UINT uFlags, UINT *pwReserved, LPSTR pszName, UINT cchMax)
{
    if (idCmd != 0)
    {
        return E_INVALIDARG;
    }

    if (uFlags == GCS_VERBW)
    {
        StringCchCopyW(reinterpret_cast<LPWSTR>(pszName), cchMax, L"ShowSpectrogram");
        return S_OK;
    }
    else if (uFlags == GCS_HELPTEXTW)
    {
        StringCchCopyW(reinterpret_cast<LPWSTR>(pszName), cchMax, 
                       L"Generate and show audio file spectrogram");
        return S_OK;
    }

    return E_INVALIDARG;
}

std::wstring SpectrogramContextMenu::GetTempPath()
{
    WCHAR tempPath[MAX_PATH];
    DWORD result = ::GetTempPathW(MAX_PATH, tempPath);
    
    if (result > 0 && result < MAX_PATH)
    {
        return std::wstring(tempPath) + L"spec.png";
    }
    
    return L"C:\\Temp\\spec.png";
}

bool SpectrogramContextMenu::FindFFmpeg(std::wstring& ffmpegPath)
{
    // Проверяем в PATH
    WCHAR buffer[MAX_PATH];
    if (SearchPathW(NULL, L"ffmpeg.exe", NULL, MAX_PATH, buffer, NULL))
    {
        ffmpegPath = buffer;
        return true;
    }
   
    return false;
}

void SpectrogramContextMenu::GenerateAndShowSpectrogram(const std::wstring& audioFile)
{
    std::wstring ffmpegPath;
    if (!FindFFmpeg(ffmpegPath))
    {
        MessageBoxW(NULL, 
                   L"ffmpeg.exe не найден. Убедитесь, что FFmpeg установлен и доступен в PATH.",
                   L"Ошибка",
                   MB_OK | MB_ICONERROR);
        return;
    }

    std::wstring outputPath = GetTempPath();
    
    // Удаляем старый файл, если существует
    DeleteFileW(outputPath.c_str());

    // Формируем командную строку для ffmpeg
    std::wstring cmdLine = L"\"" + ffmpegPath + L"\" -i \"" + audioFile + 
                          L"\" -lavfi \"showspectrumpic=s=1280x720:mode=combined:color=cool:scale=log,unsharp=7:7:1.5\" \"" + 
                          outputPath + L"\"";

    // Запускаем ffmpeg
    STARTUPINFOW si = { sizeof(si) };
    si.dwFlags = STARTF_USESHOWWINDOW;
    si.wShowWindow = SW_HIDE;
    
    PROCESS_INFORMATION pi = { 0 };

    if (CreateProcessW(NULL, 
                      const_cast<LPWSTR>(cmdLine.c_str()),
                      NULL, NULL, FALSE,
                      CREATE_NO_WINDOW,
                      NULL, NULL,
                      &si, &pi))
    {
        // Ждем завершения процесса
        WaitForSingleObject(pi.hProcess, 30000); // Таймаут 30 секунд
        
        CloseHandle(pi.hProcess);
        CloseHandle(pi.hThread);

        // Проверяем, создался ли файл
        if (PathFileExistsW(outputPath.c_str()))
        {
            ShowImage(outputPath);
        }
        else
        {
            MessageBoxW(NULL, 
                       L"Не удалось создать спектрограмму. Проверьте формат аудиофайла.",
                       L"Ошибка",
                       MB_OK | MB_ICONERROR);
        }
    }
    else
    {
        MessageBoxW(NULL, 
                   L"Не удалось запустить ffmpeg.exe",
                   L"Ошибка",
                   MB_OK | MB_ICONERROR);
    }
}

void SpectrogramContextMenu::ShowImage(const std::wstring& imagePath)
{
    // Используем стандартный просмотрщик Windows
    ShellExecuteW(NULL, L"open", imagePath.c_str(), NULL, NULL, SW_SHOW);
}

// ClassFactory реализация
ClassFactory::ClassFactory() : m_cRef(1)
{
    InterlockedIncrement(&g_cDllRef);
}

ClassFactory::~ClassFactory()
{
    InterlockedDecrement(&g_cDllRef);
}

STDMETHODIMP ClassFactory::QueryInterface(REFIID riid, void **ppvObject)
{
    if (riid == IID_IUnknown || riid == IID_IClassFactory)
    {
        *ppvObject = static_cast<IClassFactory*>(this);
        AddRef();
        return S_OK;
    }

    *ppvObject = NULL;
    return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) ClassFactory::AddRef()
{
    return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) ClassFactory::Release()
{
    ULONG cRef = InterlockedDecrement(&m_cRef);
    if (cRef == 0)
    {
        delete this;
    }
    return cRef;
}

STDMETHODIMP ClassFactory::CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppvObject)
{
    if (pUnkOuter != NULL)
    {
        return CLASS_E_NOAGGREGATION;
    }

    SpectrogramContextMenu *pMenu = new SpectrogramContextMenu();
    if (pMenu == NULL)
    {
        return E_OUTOFMEMORY;
    }

    HRESULT hr = pMenu->QueryInterface(riid, ppvObject);
    pMenu->Release();
    return hr;
}

STDMETHODIMP ClassFactory::LockServer(BOOL fLock)
{
    if (fLock)
    {
        InterlockedIncrement(&g_cDllRef);
    }
    else
    {
        InterlockedDecrement(&g_cDllRef);
    }
    return S_OK;
}

// DLL экспорты
extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
    if (dwReason == DLL_PROCESS_ATTACH)
    {
        g_hInst = hInstance;
        DisableThreadLibraryCalls(hInstance);
        InitCommonControls(); // Инициализация общих элементов управления для работы с иконками
    }
	
    return TRUE;
}

extern "C" STDAPI DllCanUnloadNow()
{
    return g_cDllRef > 0 ? S_FALSE : S_OK;
}

extern "C" STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID *ppv)
{
    if (rclsid != CLSID_SpectrogramContextMenu)
    {
        return CLASS_E_CLASSNOTAVAILABLE;
    }

    ClassFactory *pFactory = new ClassFactory();
    if (pFactory == NULL)
    {
        return E_OUTOFMEMORY;
    }

    HRESULT hr = pFactory->QueryInterface(riid, ppv);
    pFactory->Release();
    return hr;
}
