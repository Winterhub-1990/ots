#include "ContextMenuExtension.h"
#include <strsafe.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <zip.h>
#include <locale>
#include <codecvt>

// DLL instance handle
HINSTANCE g_hInst = NULL;

// Forward declarations
STDAPI DllRegisterServer();
STDAPI DllUnregisterServer();

// DLL entry point
BOOL APIENTRY DllMain(HMODULE hModule, DWORD dwReason, LPVOID lpReserved)
{
    switch (dwReason)
    {
    case DLL_PROCESS_ATTACH:
        g_hInst = hModule;
        DisableThreadLibraryCalls(hModule);
        break;
    }
    return TRUE;
}

// ContextMenuExtension implementation
ContextMenuExtension::ContextMenuExtension() : m_cRef(1), m_isDocxFile(false)
{
    DllAddRef();
}

ContextMenuExtension::~ContextMenuExtension()
{
    DllRelease();
}

STDMETHODIMP ContextMenuExtension::QueryInterface(REFIID riid, void** ppvObject)
{
    if (!ppvObject)
        return E_POINTER;

    if (IsEqualIID(riid, IID_IUnknown) ||
        IsEqualIID(riid, IID_IShellExtInit))
    {
        *ppvObject = static_cast<IShellExtInit*>(this);
    }
    else if (IsEqualIID(riid, IID_IContextMenu))
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

STDMETHODIMP_(ULONG) ContextMenuExtension::AddRef()
{
    return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) ContextMenuExtension::Release()
{
    LONG cRef = InterlockedDecrement(&m_cRef);
    if (cRef == 0)
        delete this;
    return cRef;
}

STDMETHODIMP ContextMenuExtension::Initialize(LPCITEMIDLIST pidlFolder, LPDATAOBJECT pDataObj, HKEY hProgID)
{
    if (!pDataObj)
        return E_INVALIDARG;

    FORMATETC fe = { CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
    STGMEDIUM stm;

    if (FAILED(pDataObj->GetData(&fe, &stm)))
        return E_FAIL;

    HDROP hDrop = static_cast<HDROP>(GlobalLock(stm.hGlobal));
    if (!hDrop)
    {
        ReleaseStgMedium(&stm);
        return E_FAIL;
    }

    UINT nFiles = DragQueryFile(hDrop, 0xFFFFFFFF, NULL, 0);
    m_selectedFiles.clear();

    for (UINT i = 0; i < nFiles; i++)
    {
        wchar_t szFile[MAX_PATH];
        if (DragQueryFile(hDrop, i, szFile, ARRAYSIZE(szFile)))
        {
            m_selectedFiles.push_back(szFile);
            if (IsSupportedFormat(szFile))
            {
                m_isDocxFile = true;
            }
        }
    }

    GlobalUnlock(stm.hGlobal);
    ReleaseStgMedium(&stm);

    return S_OK;
}

STDMETHODIMP ContextMenuExtension::QueryContextMenu(HMENU hMenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags)
{
    if (!m_isDocxFile)
        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, 0);

    // Create submenu
    HMENU hSubMenu = CreatePopupMenu();
    if (!hSubMenu)
        return E_FAIL;

    // Add menu items based on namespace checks
    bool showConvert = false;
    bool showExport = false;

    for (const auto& file : m_selectedFiles)
    {
        if (IsSupportedFormat(file))
        {
            if (!CheckNamespaceInDocx(file, L"http://schemas.microsoft.com/office/word/2003/wordml"))
                showConvert = true;
            if (CheckNamespaceInDocx(file, L"http://schemas.microsoft.com/office/word/2003/wordml"))
                showExport = true;
        }
    }

    UINT uFlags = MF_STRING | MF_POPUP;
    if (showConvert)
    {
        AppendMenu(hSubMenu, uFlags, idCmdFirst + MENU_CONVERT, GetLocalizedString(IDS_MENU_CONVERT).c_str());
    }
    if (showExport)
    {
        AppendMenu(hSubMenu, uFlags, idCmdFirst + MENU_EXPORT, GetLocalizedString(IDS_MENU_EXPORT).c_str());
    }
    AppendMenu(hSubMenu, uFlags, idCmdFirst + MENU_OTHER, GetLocalizedString(IDS_MENU_OTHER).c_str());

    // Add main menu item with icon
    MENUITEMINFO mii = { sizeof(mii) };
    mii.fMask = MIIM_STRING | MIIM_SUBMENU | MIIM_BITMAP;
    mii.dwTypeData = const_cast<LPWSTR>(GetLocalizedString(IDS_MENU_CONVERT).c_str());
    mii.hSubMenu = hSubMenu;
    mii.hbmpItem = LoadBitmap(g_hInst, MAKEINTRESOURCE(IDB_MENU_ICON));

    InsertMenuItem(hMenu, indexMenu, TRUE, &mii);

    return MAKE_HRESULT(SEVERITY_SUCCESS, 0, 3);
}

STDMETHODIMP ContextMenuExtension::InvokeCommand(LPCMINVOKECOMMANDINFO pici)
{
    if (!pici)
        return E_INVALIDARG;

    if (HIWORD(pici->lpVerb) != 0)
        return E_INVALIDARG;

    UINT idCmd = LOWORD(pici->lpVerb);
    std::wstring command;
    std::wstring args;

    switch (idCmd)
    {
    case MENU_CONVERT:
        command = L"convert.exe";
        args = L"-convert " + m_selectedFiles[0];
        break;
    case MENU_EXPORT:
        command = L"export.exe";
        args = L"-export " + m_selectedFiles[0];
        break;
    case MENU_OTHER:
        command = L"other.exe";
        args = L"-other " + m_selectedFiles[0];
        break;
    default:
        return E_INVALIDARG;
    }

    ExecuteCommand(command, args);
    return S_OK;
}

STDMETHODIMP ContextMenuExtension::GetCommandString(UINT_PTR idCmd, UINT uFlags, UINT* pwReserved, LPSTR pszName, UINT cchMax)
{
    if (uFlags == GCS_HELPTEXT)
    {
        std::wstring helpText;
        switch (idCmd)
        {
        case MENU_CONVERT:
            helpText = GetLocalizedString(IDS_SUBMENU_1);
            break;
        case MENU_EXPORT:
            helpText = GetLocalizedString(IDS_SUBMENU_2);
            break;
        case MENU_OTHER:
            helpText = GetLocalizedString(IDS_SUBMENU_3);
            break;
        }

        if (uFlags & GCS_UNICODE)
        {
            StringCchCopyW((LPWSTR)pszName, cchMax, helpText.c_str());
        }
        else
        {
            std::wstring_convert<std::codecvt_utf8<wchar_t>> converter;
            std::string narrow = converter.to_bytes(helpText);
            StringCchCopyA(pszName, cchMax, narrow.c_str());
        }
        return S_OK;
    }
    return E_INVALIDARG;
}

bool ContextMenuExtension::IsSupportedFormat(const std::wstring& filePath)
{
    std::wstring ext = PathFindExtension(filePath.c_str());
    for (const auto& format : SUPPORTED_FORMATS)
    {
        if (_wcsicmp(ext.c_str(), format.c_str()) == 0)
            return true;
    }
    return false;
}

bool ContextMenuExtension::CheckNamespaceInDocx(const std::wstring& filePath, const std::wstring& targetNamespace)
{
    // Implementation for checking namespaces in DOCX file
    // This is a simplified version - you'll need to implement proper ZIP and XML parsing
    struct zip* archive = zip_open(filePath.c_str(), 0, NULL);
    if (!archive)
        return false;

    bool found = false;
    struct zip_file* file;
    struct zip_stat stat;
    int numEntries = zip_get_num_entries(archive, 0);

    for (int i = 0; i < numEntries; i++)
    {
        if (zip_stat_index(archive, i, 0, &stat) == 0)
        {
            if (stat.size > 0)
            {
                file = zip_fopen_index(archive, i, 0);
                if (file)
                {
                    char* contents = new char[stat.size + 1];
                    zip_fread(file, contents, stat.size);
                    contents[stat.size] = '\0';

                    // Check for namespace in XML content
                    std::string content(contents);
                    if (content.find(targetNamespace.begin(), targetNamespace.end()) != std::string::npos)
                    {
                        found = true;
                    }

                    delete[] contents;
                    zip_fclose(file);
                }
            }
        }
    }

    zip_close(archive);
    return found;
}

void ContextMenuExtension::ExecuteCommand(const std::wstring& command, const std::wstring& args)
{
    STARTUPINFO si = { sizeof(si) };
    PROCESS_INFORMATION pi;
    std::wstring cmdLine = command + L" " + args;

    if (CreateProcess(NULL, const_cast<LPWSTR>(cmdLine.c_str()), NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi))
    {
        CloseHandle(pi.hProcess);
        CloseHandle(pi.hThread);
    }
}

std::wstring ContextMenuExtension::GetLocalizedString(UINT resourceId)
{
    wchar_t buffer[256];
    LoadString(g_hInst, resourceId, buffer, ARRAYSIZE(buffer));
    return buffer;
}

// ExplorerCommand implementation
ExplorerCommand::ExplorerCommand() : m_cRef(1), m_isDocxFile(false)
{
    DllAddRef();
}

ExplorerCommand::~ExplorerCommand()
{
    DllRelease();
}

// ... Similar implementations for ExplorerCommand methods ...

// DLL exports
STDAPI DllRegisterServer()
{
    wchar_t szModule[MAX_PATH];
    GetModuleFileName(g_hInst, szModule, ARRAYSIZE(szModule));

    // Register the COM server
    wchar_t szKey[256];
    StringCchPrintf(szKey, ARRAYSIZE(szKey), L"CLSID\\{%08lX-%04X-%04X-%02X%02X-%02X%02X%02X%02X%02X%02X}",
        CLSID_ContextMenuExtension.Data1, CLSID_ContextMenuExtension.Data2, CLSID_ContextMenuExtension.Data3,
        CLSID_ContextMenuExtension.Data4[0], CLSID_ContextMenuExtension.Data4[1],
        CLSID_ContextMenuExtension.Data4[2], CLSID_ContextMenuExtension.Data4[3],
        CLSID_ContextMenuExtension.Data4[4], CLSID_ContextMenuExtension.Data4[5],
        CLSID_ContextMenuExtension.Data4[6], CLSID_ContextMenuExtension.Data4[7]);

    HKEY hKey;
    if (RegCreateKeyEx(HKEY_CLASSES_ROOT, szKey, 0, NULL, 0, KEY_WRITE, NULL, &hKey, NULL) == ERROR_SUCCESS)
    {
        wchar_t szValue[MAX_PATH];
        StringCchPrintf(szValue, ARRAYSIZE(szValue), L"Context Menu Extension");
        RegSetValueEx(hKey, NULL, 0, REG_SZ, (BYTE*)szValue, (wcslen(szValue) + 1) * sizeof(wchar_t));

        HKEY hSubKey;
        if (RegCreateKeyEx(hKey, L"InProcServer32", 0, NULL, 0, KEY_WRITE, NULL, &hSubKey, NULL) == ERROR_SUCCESS)
        {
            RegSetValueEx(hSubKey, NULL, 0, REG_SZ, (BYTE*)szModule, (wcslen(szModule) + 1) * sizeof(wchar_t));
            RegSetValueEx(hSubKey, L"ThreadingModel", 0, REG_SZ, (BYTE*)L"Apartment", (wcslen(L"Apartment") + 1) * sizeof(wchar_t));
            RegCloseKey(hSubKey);
        }
        RegCloseKey(hKey);
    }

    // Register file types
    for (const auto& format : SUPPORTED_FORMATS)
    {
        StringCchPrintf(szKey, ARRAYSIZE(szKey), L"%s\\shellex\\ContextMenuHandlers\\ContextMenuExtension", format.c_str());
        if (RegCreateKeyEx(HKEY_CLASSES_ROOT, szKey, 0, NULL, 0, KEY_WRITE, NULL, &hKey, NULL) == ERROR_SUCCESS)
        {
            RegSetValueEx(hKey, NULL, 0, REG_SZ, (BYTE*)szModule, (wcslen(szModule) + 1) * sizeof(wchar_t));
            RegCloseKey(hKey);
        }
    }

    return S_OK;
}

STDAPI DllUnregisterServer()
{
    wchar_t szKey[256];

    // Unregister file types
    for (const auto& format : SUPPORTED_FORMATS)
    {
        StringCchPrintf(szKey, ARRAYSIZE(szKey), L"%s\\shellex\\ContextMenuHandlers\\ContextMenuExtension", format.c_str());
        RegDeleteTree(HKEY_CLASSES_ROOT, szKey);
    }

    // Unregister the COM server
    StringCchPrintf(szKey, ARRAYSIZE(szKey), L"CLSID\\{%08lX-%04X-%04X-%02X%02X-%02X%02X%02X%02X%02X%02X}",
        CLSID_ContextMenuExtension.Data1, CLSID_ContextMenuExtension.Data2, CLSID_ContextMenuExtension.Data3,
        CLSID_ContextMenuExtension.Data4[0], CLSID_ContextMenuExtension.Data4[1],
        CLSID_ContextMenuExtension.Data4[2], CLSID_ContextMenuExtension.Data4[3],
        CLSID_ContextMenuExtension.Data4[4], CLSID_ContextMenuExtension.Data4[5],
        CLSID_ContextMenuExtension.Data4[6], CLSID_ContextMenuExtension.Data4[7]);

    RegDeleteTree(HKEY_CLASSES_ROOT, szKey);

    return S_OK;
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
    if (!IsEqualCLSID(rclsid, CLSID_ContextMenuExtension))
        return CLASS_E_CLASSNOTAVAILABLE;

    ContextMenuExtension* pExt = new ContextMenuExtension();
    if (!pExt)
        return E_OUTOFMEMORY;

    HRESULT hr = pExt->QueryInterface(riid, ppv);
    pExt->Release();
    return hr;
}

STDAPI DllCanUnloadNow()
{
    return DllRefCount() == 0 ? S_OK : S_FALSE;
} 