#pragma once

#include <windows.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <string>
#include <vector>

#pragma comment(lib, "shlwapi.lib")

// GUID для нашего расширения
// {B5E2A3D1-9C4F-4E7A-8B3D-1F2E3C4D5A6B}
static const GUID CLSID_SpectrogramContextMenu = 
{ 0xb5e2a3d1, 0x9c4f, 0x4e7a, { 0x8b, 0x3d, 0x1f, 0x2e, 0x3c, 0x4d, 0x5a, 0x6b } };

class SpectrogramContextMenu : public IShellExtInit, public IContextMenu
{
public:
    SpectrogramContextMenu();
    
    // IUnknown
    STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
    STDMETHODIMP_(ULONG) AddRef();
    STDMETHODIMP_(ULONG) Release();

    // IShellExtInit
    STDMETHODIMP Initialize(LPCITEMIDLIST pidlFolder, LPDATAOBJECT pDataObj, HKEY hKeyProgID);

    // IContextMenu
    STDMETHODIMP QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags);
    STDMETHODIMP InvokeCommand(LPCMINVOKECOMMANDINFO pici);
    STDMETHODIMP GetCommandString(UINT_PTR idCmd, UINT uFlags, UINT *pwReserved, LPSTR pszName, UINT cchMax);

protected:
    ~SpectrogramContextMenu();

private:
    LONG m_cRef;
    std::vector<std::wstring> m_selectedFiles;
    
    bool IsAudioFile(const std::wstring& filename);
    void GenerateAndShowSpectrogram(const std::wstring& audioFile);
    std::wstring GetTempPath();
    bool FindFFmpeg(std::wstring& ffmpegPath);
    void ShowImage(const std::wstring& imagePath);
};

class ClassFactory : public IClassFactory
{
public:
    ClassFactory();
    
    // IUnknown
    STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
    STDMETHODIMP_(ULONG) AddRef();
    STDMETHODIMP_(ULONG) Release();

    // IClassFactory
    STDMETHODIMP CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppvObject);
    STDMETHODIMP LockServer(BOOL fLock);

protected:
    ~ClassFactory();

private:
    LONG m_cRef;
};
