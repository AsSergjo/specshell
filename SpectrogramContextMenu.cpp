#include "SpectrogramContextMenu.h"
#include "resource.h"
#include <strsafe.h>
#include <shellapi.h>
#include <commctrl.h>
#include <gdiplus.h>
#pragma comment(lib, "gdiplus.lib")

static LONG g_cDllRef = 0;
HINSTANCE g_hInst = NULL; 

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
        DestroyIcon(hIcon);
        return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, 1);
    }

    // Если не удалось добавить элемент с иконкой, пробуем без иконки
    DeleteObject(hbm);
    DestroyIcon(hIcon);
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

// Вспомогательная — получить CLSID энкодера по MIME
static int GetEncoderClsid(const WCHAR* format, CLSID* pClsid)
{
    using namespace Gdiplus;
    UINT num = 0, size = 0;
    GetImageEncodersSize(&num, &size);
    if (size == 0) return -1;

    ImageCodecInfo* pInfo = (ImageCodecInfo*)malloc(size);
    if (!pInfo) return -1;

    GetImageEncoders(num, size, pInfo);
    for (UINT i = 0; i < num; i++) {
        if (wcscmp(pInfo[i].MimeType, format) == 0) {
            *pClsid = pInfo[i].Clsid;
            free(pInfo);
            return i;
        }
    }
    free(pInfo);
    return -1;
}

// Получить sample rate через ffprobe
static bool GetAudioInfo(const std::wstring& audioFile, const std::wstring& ffmpegPath, double& sr, double& duration)
{
    sr = 44100.0;
    duration = 0.0;

    std::wstring ffprobePath = ffmpegPath;
    size_t pos = ffprobePath.rfind(L"ffmpeg.exe");
    if (pos != std::wstring::npos)
        ffprobePath.replace(pos, 10, L"ffprobe.exe");
    else
        return false;

    if (!PathFileExistsW(ffprobePath.c_str()))
        return false;

    WCHAR tempFile[MAX_PATH];
    GetTempPathW(MAX_PATH, tempFile);
    wcscat_s(tempFile, L"ffprobe_info.txt");

    // stream и format entries через двоеточие в одном -show_entries.
    // ffprobe с csv=p=0 выводит сначала sr (stream), потом duration (format),
    std::wstring cmd = L"cmd /c \"\"" + ffprobePath + 
        L"\" -v error -show_entries stream=sample_rate:format=duration "
        L"-of csv=p=0 \"" + audioFile + L"\" > \"" + tempFile + L"\"\"";

    STARTUPINFOW si = { sizeof(si) };
    si.dwFlags = STARTF_USESHOWWINDOW;
    si.wShowWindow = SW_HIDE;
    PROCESS_INFORMATION pi = {};

    if (!CreateProcessW(NULL, const_cast<LPWSTR>(cmd.c_str()),
        NULL, NULL, FALSE, CREATE_NO_WINDOW, NULL, NULL, &si, &pi))
    {
        return false;
    }

    WaitForSingleObject(pi.hProcess, 5000);
    CloseHandle(pi.hProcess);
    CloseHandle(pi.hThread);

    if (!PathFileExistsW(tempFile))
        return false;

    FILE* f = nullptr;
    if (_wfopen_s(&f, tempFile, L"r") == 0 && f) {
        char line[256];
        int lineNum = 0;
        while (fgets(line, sizeof(line), f)) {
            // Пропускаем пустые строки
            char* p = line;
            while (*p == ' ' || *p == '\r' || *p == '\n' || *p == '\t') p++;
            if (*p == '\0') continue;

            if (lineNum == 0) sscanf_s(p, "%lf", &sr);
            else if (lineNum == 1) sscanf_s(p, "%lf", &duration);
            lineNum++;
        }
        fclose(f);
    }
    DeleteFileW(tempFile);
    return true;
}

void DrawFreqLabels(const wchar_t* pngPath, double sr, double duration)
{
    using namespace Gdiplus;

    GdiplusStartupInput gsi;
    ULONG_PTR token;
    GdiplusStartup(&token, &gsi, nullptr);

    std::wstring tmpSave = std::wstring(pngPath) + L".tmp";
    bool savedOk = false;

    {
        Bitmap bmp(pngPath);
        if (bmp.GetLastStatus() != Ok) {
            GdiplusShutdown(token);
            return;
        }

        Graphics g(&bmp);
        int w = (int)bmp.GetWidth();
        int h = (int)bmp.GetHeight();

        double f_max = sr / 2.0;

        // Подбираем шаг: делим f_max на ~18, округляем до 1,2,5 * 10^n
        double rawStep = f_max / 18.0;
        double mag = pow(10.0, floor(log10(rawStep)));
        double norm = rawStep / mag;
        double niceStep;
        if      (norm < 1.5) niceStep = 1.0 * mag;
        else if (norm < 3.5) niceStep = 2.0 * mag;
        else if (norm < 7.5) niceStep = 5.0 * mag;
        else                 niceStep = 10.0 * mag;

        Font font(L"Arial", 8);
        SolidBrush brush(Color(220, 200, 200, 200));
        SolidBrush bgBrush(Color(150, 0, 0, 0));
        Pen pen(Color(90, 180, 180, 180), 1);
        pen.SetDashStyle(DashStyleDash);

        // Считаем количество тиков заранее
        int tickCount = (int)floor((f_max - 1.0) / niceStep); // последний индекс = tickCount
        for (int i = 1; i <= tickCount; i++) {
            double f = niceStep * i;
            int y = (int)round(h * (1.0 - f / f_max));
            if (y < 0 || y >= h) continue;

            g.DrawLine(&pen, 0, y, w, y);

            wchar_t label[32];
            if (i == ( tickCount - 1)) {
                // верхний тик- добавляем SR
                wchar_t srStr[16];
                if ((int)sr % 1000 == 0)
                    swprintf_s(srStr, L"  SR %.0fkHz", sr / 1000.0);
                else
                    swprintf_s(srStr, L"  SR %.1fkHz", sr / 1000.0);

                if (f >= 1000.0)
                    swprintf_s(label, L"%.0fk%s", f / 1000.0, srStr);
                else
                    swprintf_s(label, L"%.0f%s", f, srStr);
            } else {
                if (f >= 1000.0)
                    swprintf_s(label, L"%.0fk", f / 1000.0);
                else
                    swprintf_s(label, L"%.0f", f);
            }

            RectF bbox;
            g.MeasureString(label, -1, &font, PointF(0, 0), &bbox);
            g.FillRectangle(&bgBrush, 2, y - 12, (int)bbox.Width + 4, 12);
            g.DrawString(label, -1, &font, PointF(3, (float)(y - 13)), &brush);
        }

       if (duration > 0.0) {
            int divisions = 16;
            double timeStep = duration / divisions;

            for (int i = 0; i <= divisions; i++) {
                double t = i * timeStep;
                int x = (int)round((double)w * i / divisions);
                if (x < 0 || x > w) continue;

                g.DrawLine(&pen, x, 0, x, h - 15);

                wchar_t label[32];
                int min = (int)(t / 60.0);
                int sec = (int)(t) % 60;
                swprintf_s(label, L"%d:%02d", min, sec);

                RectF bbox;
                g.MeasureString(label, -1, &font, PointF(0, 0), &bbox);
                int labelX = (i == 0) ? x : x - (int)bbox.Width / 2;
                if (labelX < 0) labelX = 0;
                if (labelX + (int)bbox.Width > w) labelX = w - (int)bbox.Width;

                g.FillRectangle(&bgBrush, labelX, h - 14, (int)bbox.Width + 4, 12);
                g.DrawString(label, -1, &font, PointF((float)(labelX + 2), (float)(h - 13)), &brush);
            }
        }

        CLSID clsid;
        if (GetEncoderClsid(L"image/png", &clsid) >= 0)
            savedOk = (bmp.Save(tmpSave.c_str(), &clsid) == Ok);
    }

    GdiplusShutdown(token);

    if (savedOk)
        MoveFileExW(tmpSave.c_str(), pngPath, MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH);
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
                          L"\" -lavfi \"showspectrumpic=s=" +
                          std::to_wstring(SPECTROGRAM_WIDTH) + L"x" +
                          std::to_wstring(SPECTROGRAM_HEIGHT) +
                          L":mode=" + std::wstring(SPECTROGRAM_MODE) +
                          L":color=" + std::wstring(SPECTROGRAM_COLOR) +
                          L":scale=" + std::wstring(SPECTROGRAM_SCALE) +
                          L":legend=0,unsharp=" + std::wstring(SPECTROGRAM_UNSHARP) +
                          L"\" \"" + outputPath + L"\"";

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
        WaitForSingleObject(pi.hProcess, FFMPEG_TIMEOUT);

        CloseHandle(pi.hProcess);
        CloseHandle(pi.hThread);

        if (PathFileExistsW(outputPath.c_str()))
        {   
            double sr, duration;
            GetAudioInfo(audioFile, ffmpegPath, sr, duration);
            DrawFreqLabels(outputPath.c_str(), sr, duration);
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
    HINSTANCE result = ShellExecuteW(NULL, L"open", imagePath.c_str(), NULL, NULL, SW_SHOW);
    if ((INT_PTR)result <= 32)
    {
        MessageBoxW(NULL,
                   L"Не удалось открыть изображение. Попробуйте открыть файл вручную.",
                   L"Ошибка",
                   MB_OK | MB_ICONERROR);
    }
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
