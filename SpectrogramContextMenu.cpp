#include "SpectrogramContextMenu.h"
#include "resource.h"
#include <strsafe.h>
#include <shellapi.h>
#include <commctrl.h>
#include <gdiplus.h>
#include <cstdlib>
#include <vector>
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

// Вспомогательная - получить CLSID энкодера по MIME
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

static void TrimAscii(std::string& s)
{
    size_t start = 0;
    while (start < s.size() && (s[start] == ' ' || s[start] == '\t'))
        ++start;

    size_t end = s.size();
    while (end > start && (s[end - 1] == ' ' || s[end - 1] == '\t'))
        --end;

    s = s.substr(start, end - start);
}

static bool TryParseDouble(const std::string& s, double& value)
{
    char* end = nullptr;
    value = std::strtod(s.c_str(), &end);
    if (end == s.c_str())
        return false;

    while (*end == ' ' || *end == '\t')
        ++end;

    return *end == '\0';
}

static bool TryParseInt(const std::string& s, int& value)
{
    char* end = nullptr;
    long parsed = std::strtol(s.c_str(), &end, 10);
    if (end == s.c_str())
        return false;

    while (*end == ' ' || *end == '\t')
        ++end;

    if (*end != '\0')
        return false;

    value = (int)parsed;
    return true;
}

// Вспомогательная: char* -> wstring (UTF-8 / ANSI)
static std::wstring MBtoWS(const std::string& s)
{
    if (s.empty()) return L"";

    UINT codePage = CP_UTF8;
    int n = MultiByteToWideChar(codePage, MB_ERR_INVALID_CHARS, s.c_str(), -1, nullptr, 0);
    if (n <= 0) {
        codePage = CP_ACP;
        n = MultiByteToWideChar(codePage, 0, s.c_str(), -1, nullptr, 0);
    }
    if (n <= 0) return L"";

    std::vector<wchar_t> buf(n);
    MultiByteToWideChar(codePage, 0, s.c_str(), -1, buf.data(), n);
    return buf.data();
}

// Получить sample rate, duration, codec, bitrate, format через ffprobe
static bool GetAudioInfo(const std::wstring& audioFile, const std::wstring& ffmpegPath,
    double& sr, double& duration,
    std::wstring& codec, int& bitrate, std::wstring& format)
{
    sr       = 44100.0;
    duration = 0.0;
    codec    = L"";
    bitrate  = 0;
    format   = L"";

    std::wstring ffprobePath = ffmpegPath;
    size_t pos = ffprobePath.rfind(L"ffmpeg.exe");
    if (pos != std::wstring::npos)
        ffprobePath.replace(pos, 10, L"ffprobe.exe");
    else
        return false;

    if (!PathFileExistsW(ffprobePath.c_str()))
        return false;

    SECURITY_ATTRIBUTES sa = {};
    sa.nLength = sizeof(sa);
    sa.bInheritHandle = TRUE;

    HANDLE hRead = NULL;
    HANDLE hWrite = NULL;
    if (!CreatePipe(&hRead, &hWrite, &sa, 0))
        return false;

    SetHandleInformation(hRead, HANDLE_FLAG_INHERIT, 0);

    HANDLE hNull = CreateFileW(L"NUL", GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE,
        NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
    if (hNull == INVALID_HANDLE_VALUE) {
        CloseHandle(hRead);
        CloseHandle(hWrite);
        return false;
    }

    // Именованный вывод без section wrappers:
    //   sample_rate=48000
    //   codec_name=mp3
    //   duration=215.420000
    //   bit_rate=320000
    //   format_name=mp3,mpegaudio
    std::wstring cmd = L"\"" + ffprobePath +
        L"\" -v error -select_streams a:0"
        L" -show_entries stream=sample_rate,codec_name:format=duration,bit_rate,format_name"
        L" -of default=noprint_wrappers=1 \"" + audioFile + L"\"";

    STARTUPINFOW si = { sizeof(si) };
    si.dwFlags = STARTF_USESHOWWINDOW | STARTF_USESTDHANDLES;
    si.wShowWindow = SW_HIDE;
    si.hStdOutput = hWrite;
    si.hStdError = hNull;
    si.hStdInput = GetStdHandle(STD_INPUT_HANDLE);

    PROCESS_INFORMATION pi = {};
    BOOL ok = CreateProcessW(NULL, const_cast<LPWSTR>(cmd.c_str()),
        NULL, NULL, TRUE, CREATE_NO_WINDOW, NULL, NULL, &si, &pi);

    CloseHandle(hWrite);
    CloseHandle(hNull);

    if (!ok) {
        CloseHandle(hRead);
        return false;
    }

    std::string output;
    char buffer[512];
    for (;;) {
        DWORD bytesRead = 0;
        BOOL readOk = ReadFile(hRead, buffer, sizeof(buffer), &bytesRead, NULL);
        if (!readOk || bytesRead == 0)
            break;
        output.append(buffer, bytesRead);
    }

    WaitForSingleObject(pi.hProcess, 5000);
    CloseHandle(hRead);
    CloseHandle(pi.hProcess);
    CloseHandle(pi.hThread);

    size_t start = 0;
    while (start <= output.size()) {
        size_t end = output.find('\n', start);
        std::string line = (end == std::string::npos)
            ? output.substr(start)
            : output.substr(start, end - start);
        if (!line.empty() && line.back() == '\r')
            line.pop_back();
        TrimAscii(line);
        if (!line.empty()) {
            size_t sep = line.find('=');
            if (sep != std::string::npos) {
                std::string key = line.substr(0, sep);
                std::string value = line.substr(sep + 1);
                TrimAscii(key);
                TrimAscii(value);

                if (key == "sample_rate") {
                    double parsedSr = 0.0;
                    if (TryParseDouble(value, parsedSr) && parsedSr > 0.0)
                        sr = parsedSr;
                }
                else if (key == "codec_name") {
                    codec = MBtoWS(value);
                }
                else if (key == "duration") {
                    double parsedDuration = 0.0;
                    if (TryParseDouble(value, parsedDuration) && parsedDuration >= 0.0)
                        duration = parsedDuration;
                }
                else if (key == "bit_rate") {
                    int parsedBitrate = 0;
                    if (TryParseInt(value, parsedBitrate) && parsedBitrate > 0)
                        bitrate = parsedBitrate;
                }
                else if (key == "format_name") {
                    format = MBtoWS(value);
                }
            }
        }
        if (end == std::string::npos)
            break;
        start = end + 1;
    }

    return true;
}

void DrawFreqLabels(const wchar_t* pngPath, double sr, double duration,
    const std::wstring& codec, int bitrate, const std::wstring& format,
    const std::wstring& filename)
{
    using namespace Gdiplus;

    GdiplusStartupInput gsi;
    ULONG_PTR token;
    GdiplusStartup(&token, &gsi, nullptr);

    std::wstring tmpSave = std::wstring(pngPath) + L".tmp";
    DeleteFileW(tmpSave.c_str());
    bool savedOk = false;

    {
        // Загружаем оригинальную спектрограмму
        Bitmap src(pngPath);
        if (src.GetLastStatus() != Gdiplus::Ok) {
            GdiplusShutdown(token);
            return;
        }

        int srcW = (int)src.GetWidth();
        int srcH = (int)src.GetHeight();

        if (sr <= 1000.0)
            sr = 44100.0;
        if (duration < 0.0)
            duration = 0.0;

        // Поля вокруг спектрограммы (в пикселях)
        const int LEFT      = 54;  // для меток частот (ось Y)
        const int RIGHT     = 6;
        const int INFO_H    = 22;  // высота поля с информацией над спектрограммой
        const int TOP       = INFO_H + 6; // = 28
        const int BOTTOM    = 22;  // для меток времени (ось X)

        int totalW = LEFT + srcW + RIGHT;
        int totalH = TOP  + srcH + BOTTOM;

        // Новый bitmap - спектрограмма + поля
        Bitmap bmp(totalW, totalH, PixelFormat32bppARGB);
        Graphics g(&bmp);
        g.SetSmoothingMode(SmoothingModeNone);
        g.SetTextRenderingHint(TextRenderingHintSingleBitPerPixelGridFit);

        // Фон - чёрный
        g.Clear(Color(255, 0, 0, 0));

        // Вписываем оригинальную спектрограмму в центральную область
        g.DrawImage(&src, LEFT, TOP, srcW, srcH);

        // Тонкая рамка вокруг спектрограммы (1 px, серая)
        Pen borderPen(Color(255, 100, 100, 100), 1.0f);
        g.DrawRectangle(&borderPen, LEFT, TOP, srcW - 1, srcH - 1);

        // ── Информационное поле над спектрограммой ──────────────────────────
        {
            // Формируем строку: SR xx.xk  |  xxx kbps  |  format  |  codec  |  filename
            wchar_t srPart[32];
            if ((int)sr % 1000 == 0)
                swprintf_s(srPart, L"SR %.0fk", sr / 1000.0);
            else
                swprintf_s(srPart, L"SR %.1fk", sr / 1000.0);

            wchar_t brPart[32];
            if (bitrate > 0)
                swprintf_s(brPart, L"%d kbps", bitrate / 1000);
            else
                wcscpy_s(brPart, L"? kbps");

            // format_name может быть составным ("mp3,mpegaudio") — берём первый токен
            std::wstring fmtShort = format;
            size_t comma = fmtShort.find(L',');
            if (comma != std::wstring::npos) fmtShort = fmtShort.substr(0, comma);

            wchar_t infoLine[512];
            swprintf_s(infoLine,
                L"%s  |  %s  |  %s  |  %s  |  %s",
                srPart,
                brPart,
                fmtShort.empty()  ? L"?" : fmtShort.c_str(),
                codec.empty()     ? L"?" : codec.c_str(),
                filename.empty()  ? L"?" : filename.c_str());

            Font      infoFont(L"Arial", 9, FontStyleRegular);
            SolidBrush infoBrush(Color(255, 220, 220, 220));

            // Визуально центрируем текст во всей верхней зоне до спектрограммы (TOP = INFO_H + 6)
            RectF bbox;
            g.MeasureString(infoLine, -1, &infoFont, PointF(0, 0), &bbox);
            float textY = ((float)TOP - bbox.Height) * 0.5f;
            if (textY < 1.0f) textY = 1.0f;

            // Отступ слева совпадает с LEFT (под метками оси Y)
            g.DrawString(infoLine, -1, &infoFont,
                         PointF((float)LEFT, textY), &infoBrush);
        }
        // ────────────────────────────────────────────────────────────────────

        // Параметры осей
        double f_max = sr / 2.0;
        double rawStep = f_max / 18.0;
        double mag = pow(10.0, floor(log10(rawStep)));
        double norm = rawStep / mag;
        double niceStep;
        if      (norm < 1.5) niceStep = 1.0 * mag;
        else if (norm < 3.5) niceStep = 2.0 * mag;
        else if (norm < 7.5) niceStep = 5.0 * mag;
        else                 niceStep = 10.0 * mag;

        Font      font(L"Arial", 8);
        SolidBrush brush(Color(255, 200, 200, 200));
        Pen       tickPen(Color(255, 130, 130, 130), 1.0f);

        int tickCount = (int)floor((f_max - 1.0) / niceStep);

        // Ось Y: метки частот слева от спектрограммы 
        // тики - только числовые метки частот
        for (int i = 1; i <= tickCount; i++) {
            double f = niceStep * i;
            int y = TOP + (int)round(srcH * (1.0 - f / f_max));
            if (y < TOP || y > TOP + srcH) continue;

            // Тик-штрих на левой границе (4 px наружу)
            g.DrawLine(&tickPen, LEFT - 4, y, LEFT, y);

            wchar_t label[32];
            if (f >= 1000.0)
                swprintf_s(label, L"%.0fk", f / 1000.0);
            else
                swprintf_s(label, L"%.0f", f);

            RectF bbox;
            g.MeasureString(label, -1, &font, PointF(0, 0), &bbox);

            float labelX = 24.0f; //смещение по оси х
            float labelY = (float)y - bbox.Height * 0.5f;
            g.DrawString(label, -1, &font, PointF(labelX, labelY), &brush);
        }

        // Ось X: метки времени снизу от спектрограммы
        if (duration > 0.0) {
            const int divisions = 16;
            double timeStep = duration / divisions;

            for (int i = 0; i <= divisions; i++) {
                double t = i * timeStep;
                int x = LEFT + (int)round((double)srcW * i / divisions);
                if (x < LEFT || x > LEFT + srcW) continue;

                // Тик-штрих на нижней границе (4 px вниз)
                g.DrawLine(&tickPen, x, TOP + srcH, x, TOP + srcH + 4);

                wchar_t label[32];
                int min = (int)(t / 60.0);
                int sec = (int)(t) % 60;
                swprintf_s(label, L"%d:%02d", min, sec);

                RectF bbox;
                g.MeasureString(label, -1, &font, PointF(0, 0), &bbox);

                // По центру тика, с ограничением по краям
                float labelX = (float)x - bbox.Width * 0.5f;
                if (labelX < 0.0f) labelX = 0.0f;
                if (labelX + bbox.Width > (float)totalW)
                    labelX = (float)totalW - bbox.Width;

                float labelY = (float)(TOP + srcH + 6);
                g.DrawString(label, -1, &font, PointF(labelX, labelY), &brush);
            }
        }

        CLSID clsid;
        if (GetEncoderClsid(L"image/png", &clsid) >= 0)
            savedOk = (bmp.Save(tmpSave.c_str(), &clsid) == Gdiplus::Ok);
    }

    GdiplusShutdown(token);

    if (savedOk)
    {
        if (!MoveFileExW(tmpSave.c_str(), pngPath, MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH))
            DeleteFileW(tmpSave.c_str());
    }
    else
    {
        DeleteFileW(tmpSave.c_str());
    }
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
            std::wstring codec, format;
            int bitrate = 0;
            GetAudioInfo(audioFile, ffmpegPath, sr, duration, codec, bitrate, format);

            // Только имя файла (без пути)
            const wchar_t* fname = PathFindFileNameW(audioFile.c_str());
            std::wstring filename = fname ? fname : audioFile;

            DrawFreqLabels(outputPath.c_str(), sr, duration, codec, bitrate, format, filename);
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
