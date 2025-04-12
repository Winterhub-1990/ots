#include "ContextMenuExtension.h"
#include "resource.h"
#include <strsafe.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <windows.h>
#include <string>
#include <vector>
#include <iostream>
#include <sstream>
#include <algorithm>
#include <thread>
#include <atomic>
#include <locale>
#include <codecvt>
#include <filesystem>
#include <fstream>
#include <zip.h>
#include <tinyxml2.h>
#include <new>
#include <ShlGuid.h>    // For shell GUIDs
#include <ShObjIdl.h>   // For IShellItemArray
#include <shlwapi.h>    // For QITAB and QISearch
#include <objbase.h>    // For COM interfaces

// Define supported formats
const std::vector<std::wstring> SUPPORTED_FORMATS = {
    L".docx", L".docm", L".dotx", L".dotm", L".xml", L".xin"
};

// Menu command IDs
enum MenuCommands
{
    MENU_CONVERT = 0,
    MENU_EXPORT,
    MENU_OTHER
};

// Global variables
HINSTANCE g_hInst = NULL;
long g_cDllRef = 0;

// Helper function to convert wstring to string (Windows-specific)
std::string WStringToString(const std::wstring& wstr) {
    if (wstr.empty()) return std::string();
    int size_needed = WideCharToMultiByte(CP_UTF8, 0, &wstr[0], (int)wstr.size(), NULL, 0, NULL, NULL);
    std::string strTo(size_needed, 0);
    WideCharToMultiByte(CP_UTF8, 0, &wstr[0], (int)wstr.size(), &strTo[0], size_needed, NULL, NULL);
    return strTo;
}

// Helper function to convert string to wstring (Windows-specific)
std::wstring StringToWString(const std::string& str) {
    if (str.empty()) return std::wstring();
    int size_needed = MultiByteToWideChar(CP_UTF8, 0, &str[0], (int)str.size(), NULL, 0);
    std::wstring wstrTo(size_needed, 0);
    MultiByteToWideChar(CP_UTF8, 0, &str[0], (int)str.size(), &wstrTo[0], size_needed);
    return wstrTo;
}

ContextMenuExtension::ContextMenuExtension() : m_cRef(1), m_hasTargetNamespace(false) {
    InterlockedIncrement(&g_cDllRef);
}

ContextMenuExtension::~ContextMenuExtension() {
    InterlockedDecrement(&g_cDllRef);
}

// IUnknown methods
IFACEMETHODIMP ContextMenuExtension::QueryInterface(REFIID riid, void** ppv) {
    *ppv = NULL;

    // Check for known interfaces
    if (IsEqualIID(riid, IID_IUnknown)) {
        *ppv = static_cast<IExplorerCommand*>(this);
    }
    else if (IsEqualIID(riid, IID_IExplorerCommand)) {
        *ppv = static_cast<IExplorerCommand*>(this);
    }
    else if (IsEqualIID(riid, IID_IContextMenu)) {
        *ppv = static_cast<IContextMenu*>(this);
    }
    else if (IsEqualIID(riid, IID_IContextMenu2)) {
        *ppv = static_cast<IContextMenu2*>(this);
    }
    else if (IsEqualIID(riid, IID_IContextMenu3)) {
        *ppv = static_cast<IContextMenu3*>(this);
    }
    else if (IsEqualIID(riid, IID_IShellExtInit)) {
        *ppv = static_cast<IShellExtInit*>(this);
    }
    else {
        return E_NOINTERFACE;
    }

    AddRef();
    return S_OK;
}

IFACEMETHODIMP_(ULONG) ContextMenuExtension::AddRef() {
    return InterlockedIncrement(&m_cRef);
}

IFACEMETHODIMP_(ULONG) ContextMenuExtension::Release() {
    ULONG cRef = InterlockedDecrement(&m_cRef);
    if (cRef == 0) {
        delete this;
    }
    return cRef;
}

// IExplorerCommand methods
IFACEMETHODIMP ContextMenuExtension::GetTitle(IShellItemArray* psiItemArray, LPWSTR* ppszName) {
    *ppszName = nullptr;
    WCHAR szTitle[50];
    
    if (LoadStringW(g_hInst, IDS_MENU_TITLE, szTitle, ARRAYSIZE(szTitle)) == 0) {
        return E_FAIL;
    }

    size_t len = wcslen(szTitle) + 1;
    *ppszName = static_cast<LPWSTR>(CoTaskMemAlloc(len * sizeof(WCHAR)));
    if (*ppszName == nullptr) {
        return E_OUTOFMEMORY;
    }

    wcscpy_s(*ppszName, len, szTitle);
    return S_OK;
}

IFACEMETHODIMP ContextMenuExtension::GetIcon(IShellItemArray* psiItemArray, LPWSTR* ppszIcon) {
    *ppszIcon = nullptr;
    WCHAR szModulePath[MAX_PATH];
    
    if (GetModuleFileNameW(g_hInst, szModulePath, ARRAYSIZE(szModulePath)) == 0) {
        return E_FAIL;
    }

    // Format the icon resource path
    std::wstring iconPath = szModulePath;
    iconPath += L",-";
    iconPath += std::to_wstring(IDI_MENU_ICON);

    size_t len = iconPath.length() + 1;
    *ppszIcon = static_cast<LPWSTR>(CoTaskMemAlloc(len * sizeof(WCHAR)));
    if (*ppszIcon == nullptr) {
        return E_OUTOFMEMORY;
    }

    wcscpy_s(*ppszIcon, len, iconPath.c_str());
    return S_OK;
}

IFACEMETHODIMP ContextMenuExtension::GetState(IShellItemArray* psiItemArray, BOOL fOkToBeSlow, EXPCMDSTATE* pCmdState) {
    *pCmdState = ECS_ENABLED;
    
    if (psiItemArray) {
        DWORD count;
        psiItemArray->GetCount(&count);
        if (count == 1) {
            IShellItem* psi;
            psiItemArray->GetItemAt(0, &psi);
            LPWSTR pszPath;
            psi->GetDisplayName(SIGDN_FILESYSPATH, &pszPath);
            m_selectedFile = pszPath;
            CoTaskMemFree(pszPath);
            psi->Release();

            if (CheckFileExtension(m_selectedFile)) {
                m_hasTargetNamespace = CheckTargetNamespace(m_selectedFile);
                return S_OK;
            }
        }
    }
    
    *pCmdState = ECS_HIDDEN;
    return S_OK;
}

bool ContextMenuExtension::CheckFileExtension(const std::wstring& filePath) {
    std::wstring ext = std::filesystem::path(filePath).extension();
    return std::find(SUPPORTED_FORMATS.begin(), SUPPORTED_FORMATS.end(), ext) != SUPPORTED_FORMATS.end();
}

bool ContextMenuExtension::CheckTargetNamespace(const std::wstring& filePath) {
    std::string path = WStringToString(filePath);
    struct zip* archive = zip_open(path.c_str(), 0, nullptr);
    if (!archive) return false;

    bool found = false;
    zip_int64_t num_entries = zip_get_num_entries(archive, 0);
    
    for (zip_int64_t i = 0; i < num_entries; i++) {
        const char* name = zip_get_name(archive, i, 0);
        if (strstr(name, "customXML/itemProps") && strstr(name, ".xml")) {
            struct zip_file* file = zip_fopen_index(archive, i, 0);
            if (file) {
                char buffer[4096];
                zip_int64_t len;
                std::string content;
                
                while ((len = zip_fread(file, buffer, sizeof(buffer))) > 0) {
                    content.append(buffer, len);
                }
                
                zip_fclose(file);
                
                // Parse XML using tinyxml2
                tinyxml2::XMLDocument doc;
                if (doc.Parse(content.c_str()) == tinyxml2::XML_SUCCESS) {
                    auto element = doc.FirstChildElement("datastoreItem")
                                    ->FirstChildElement("schemaRefs")
                                    ->FirstChildElement("schemaRef");
                    
                    if (element) {
                        const char* uri = element->Attribute("uri");
                        if (uri && strcmp(uri, WStringToString(TARGET_NAMESPACE).c_str()) == 0) {
                            found = true;
                            break;
                        }
                    }
                }
            }
        }
    }
    
    zip_close(archive);
    return found;
}

IFACEMETHODIMP ContextMenuExtension::Invoke(IShellItemArray* psiItemArray, IBindCtx* pbc) {
    // Implementation for command invocation
    return S_OK;
}

// IContextMenu methods
IFACEMETHODIMP ContextMenuExtension::QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags) {
    if (uFlags & CMF_DEFAULTONLY) return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, 0);

    UINT idCmd = idCmdFirst;
    HMENU hSubmenu = CreatePopupMenu();
    
    if (m_hasTargetNamespace) {
        // Show Export menu
        AppendMenuW(hSubmenu, MF_STRING, idCmd++, GetLocalizedString(IDS_MENU_EXPORT).c_str());
    } else {
        // Show Convert menu
        AppendMenuW(hSubmenu, MF_STRING, idCmd++, GetLocalizedString(IDS_MENU_CONVERT).c_str());
    }
    
    // Add submenus
    AppendMenuW(hSubmenu, MF_STRING, idCmd++, GetLocalizedString(IDS_SUBMENU_1).c_str());
    AppendMenuW(hSubmenu, MF_STRING, idCmd++, GetLocalizedString(IDS_SUBMENU_2).c_str());
    AppendMenuW(hSubmenu, MF_STRING, idCmd++, GetLocalizedString(IDS_SUBMENU_3).c_str());
    
    // Insert the submenu into the context menu
    MENUITEMINFO mii = { sizeof(mii) };
    mii.fMask = MIIM_SUBMENU | MIIM_STRING | MIIM_ID;
    mii.wID = idCmd++;
    mii.hSubMenu = hSubmenu;
    mii.dwTypeData = const_cast<LPWSTR>(GetLocalizedString(IDS_MENU_TITLE).c_str());
    
    InsertMenuItem(hmenu, indexMenu, TRUE, &mii);
    
    return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, idCmd - idCmdFirst);
}

std::wstring ContextMenuExtension::GetLocalizedString(UINT resourceId) {
    WCHAR szBuffer[100];
    LoadStringW(g_hInst, resourceId, szBuffer, ARRAYSIZE(szBuffer));
    return std::wstring(szBuffer);
}

// DLL exports
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppv) {
    *ppv = NULL;
    
    if (IsEqualCLSID(CLSID_ContextMenuExtension, rclsid)) {
        ContextMenuExtensionFactory* pFactory = new ContextMenuExtensionFactory();
        if (pFactory) {
            return pFactory->QueryInterface(riid, ppv);
        }
    }
    
    return CLASS_E_CLASSNOTAVAILABLE;
}

STDAPI DllCanUnloadNow() {
    return g_cDllRef == 0 ? S_OK : S_FALSE;
}

STDAPI DllRegisterServer() {
    HKEY hKey;
    WCHAR szModule[MAX_PATH];
    GetModuleFileNameW(g_hInst, szModule, ARRAYSIZE(szModule));
    
    // Register COM server
    if (RegCreateKeyExW(HKEY_CLASSES_ROOT, L"CLSID\\{8943D31D-6EEB-4542-B2D3-666D1A089F0C}", 0, NULL, 
        REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL) == ERROR_SUCCESS) {
        RegSetValueExW(hKey, NULL, 0, REG_SZ, (BYTE*)L"Context Menu Extension", 
            sizeof(L"Context Menu Extension"));
        RegCloseKey(hKey);
        
        if (RegCreateKeyExW(HKEY_CLASSES_ROOT, 
            L"CLSID\\{8943D31D-6EEB-4542-B2D3-666D1A089F0C}\\InProcServer32", 0, NULL,
            REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL) == ERROR_SUCCESS) {
            size_t len = wcslen(szModule);
            if (len > MAXDWORD) {
                return E_FAIL;  // Path too long
            }
            RegSetValueExW(hKey, NULL, 0, REG_SZ, (BYTE*)szModule, 
                static_cast<DWORD>((len + 1) * sizeof(WCHAR)));
            RegSetValueExW(hKey, L"ThreadingModel", 0, REG_SZ, 
                (BYTE*)L"Apartment", sizeof(L"Apartment"));
            RegCloseKey(hKey);
        }
    }
    
    // Register for file types
    for (const auto& ext : SUPPORTED_FORMATS) {
        std::wstring keyPath = ext + L"\\shellex\\ContextMenuHandlers\\ContextMenuExtension";
        if (RegCreateKeyExW(HKEY_CLASSES_ROOT, keyPath.c_str(), 0, NULL,
            REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL) == ERROR_SUCCESS) {
            RegSetValueExW(hKey, NULL, 0, REG_SZ, (BYTE*)L"{8943D31D-6EEB-4542-B2D3-666D1A089F0C}",
                sizeof(L"{8943D31D-6EEB-4542-B2D3-666D1A089F0C}"));
            RegCloseKey(hKey);
        }
    }
    
    return S_OK;
}

STDAPI DllUnregisterServer() {
    // Unregister COM server
    RegDeleteTreeW(HKEY_CLASSES_ROOT, L"CLSID\\{8943D31D-6EEB-4542-B2D3-666D1A089F0C}");
    
    // Unregister file types
    for (const auto& ext : SUPPORTED_FORMATS) {
        std::wstring keyPath = ext + L"\\shellex\\ContextMenuHandlers\\ContextMenuExtension";
        RegDeleteTreeW(HKEY_CLASSES_ROOT, keyPath.c_str());
    }
    
    return S_OK;
}

// DLL entry point
BOOL APIENTRY DllMain(HMODULE hModule, DWORD reason, LPVOID reserved) {
    switch (reason) {
        case DLL_PROCESS_ATTACH:
            g_hInst = hModule;
            DisableThreadLibraryCalls(hModule);
            break;
    }
    return TRUE;
}

// Add missing IExplorerCommand implementations
IFACEMETHODIMP ContextMenuExtension::GetToolTip(IShellItemArray* psiItemArray, LPWSTR* ppszInfotip) {
    *ppszInfotip = nullptr;
    return E_NOTIMPL;
}

IFACEMETHODIMP ContextMenuExtension::GetCanonicalName(GUID* pguidCommandName) {
    *pguidCommandName = GUID_NULL;
    return E_NOTIMPL;
}

IFACEMETHODIMP ContextMenuExtension::GetFlags(EXPCMDFLAGS* pFlags) {
    *pFlags = ECF_DEFAULT;
    return S_OK;
}

IFACEMETHODIMP ContextMenuExtension::EnumSubCommands(IEnumExplorerCommand** ppEnum) {
    *ppEnum = nullptr;
    return E_NOTIMPL;
}

// Add missing IContextMenu implementations
IFACEMETHODIMP ContextMenuExtension::GetCommandString(UINT_PTR idCmd, UINT uFlags, UINT* pwReserved, LPSTR pszName, UINT cchMax) {
    return E_NOTIMPL;
}

IFACEMETHODIMP ContextMenuExtension::HandleMenuMsg(UINT uMsg, WPARAM wParam, LPARAM lParam) {
    return HandleMenuMsg2(uMsg, wParam, lParam, nullptr);
}

IFACEMETHODIMP ContextMenuExtension::HandleMenuMsg2(UINT uMsg, WPARAM wParam, LPARAM lParam, LRESULT* plResult) {
    if (plResult)
        *plResult = 0;
    return S_OK;
}

// Add missing IContextMenu::InvokeCommand implementation
IFACEMETHODIMP ContextMenuExtension::InvokeCommand(LPCMINVOKECOMMANDINFO pici) {
    if (HIWORD(pici->lpVerb))
        return E_INVALIDARG;

    UINT idCmd = LOWORD(pici->lpVerb);
    if (idCmd >= CMD_CONVERT && idCmd <= CMD_COMMAND3) {
        ExecuteCommand(static_cast<CommandId>(idCmd), m_selectedFile);
        return S_OK;
    }

    return E_INVALIDARG;
}

// Add ExecuteCommand implementation
void ContextMenuExtension::ExecuteCommand(CommandId cmdId, const std::wstring& filePath) {
    switch (cmdId) {
        case CMD_CONVERT:
            // Implement conversion logic
            break;
        case CMD_EXPORT:
            // Implement export logic
            break;
        case CMD_COMMAND1:
        case CMD_COMMAND2:
        case CMD_COMMAND3:
            // Implement other commands
            break;
    }
}

// Factory class implementation
ContextMenuExtensionFactory::ContextMenuExtensionFactory() : m_cRef(1) {
    InterlockedIncrement(&g_cDllRef);
}

ContextMenuExtensionFactory::~ContextMenuExtensionFactory() {
    InterlockedDecrement(&g_cDllRef);
}

IFACEMETHODIMP ContextMenuExtensionFactory::QueryInterface(REFIID riid, void** ppv) {
    *ppv = NULL;

    if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IClassFactory)) {
        *ppv = static_cast<IClassFactory*>(this);
        AddRef();
        return S_OK;
    }

    return E_NOINTERFACE;
}

IFACEMETHODIMP_(ULONG) ContextMenuExtensionFactory::AddRef() {
    return InterlockedIncrement(&m_cRef);
}

IFACEMETHODIMP_(ULONG) ContextMenuExtensionFactory::Release() {
    ULONG cRef = InterlockedDecrement(&m_cRef);
    if (cRef == 0) {
        delete this;
    }
    return cRef;
}

IFACEMETHODIMP ContextMenuExtensionFactory::CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppv) {
    *ppv = NULL;

    // The shell extension doesn't support aggregation
    if (pUnkOuter != NULL) {
        return CLASS_E_NOAGGREGATION;
    }

    ContextMenuExtension *pExt = new (std::nothrow) ContextMenuExtension();
    if (pExt == NULL) {
        return E_OUTOFMEMORY;
    }

    HRESULT hr = pExt->QueryInterface(riid, ppv);
    pExt->Release();

    return hr;
}

IFACEMETHODIMP ContextMenuExtensionFactory::LockServer(BOOL fLock) {
    if (fLock) {
        InterlockedIncrement(&g_cDllRef);
    } else {
        InterlockedDecrement(&g_cDllRef);
    }
    return S_OK;
}

// Add IShellExtInit implementation
IFACEMETHODIMP ContextMenuExtension::Initialize(PCIDLIST_ABSOLUTE pidlFolder, IDataObject* pdtobj, HKEY hkeyProgID) {
    if (pdtobj == nullptr) {
        return E_INVALIDARG;
    }

    FORMATETC fe = { CF_HDROP, nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
    STGMEDIUM stm;

    if (SUCCEEDED(pdtobj->GetData(&fe, &stm))) {
        HDROP hDrop = static_cast<HDROP>(stm.hGlobal);
        UINT cFiles = DragQueryFileW(hDrop, 0xFFFFFFFF, nullptr, 0);
        
        if (cFiles == 1) {  // We only handle single selection
            WCHAR szFile[MAX_PATH];
            if (DragQueryFileW(hDrop, 0, szFile, ARRAYSIZE(szFile))) {
                m_selectedFile = szFile;
            }
        }
        
        ReleaseStgMedium(&stm);
        return S_OK;
    }

    return E_FAIL;
} 