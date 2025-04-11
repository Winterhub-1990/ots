#pragma once

#include <windows.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <string>
#include <vector>
#include <memory>
#include <zip.h>

// Forward declarations
class ContextMenuExtension;
class ExplorerCommand;

// {YOUR-GUID-HERE} - Replace with your own GUID
const CLSID CLSID_ContextMenuExtension = { 0x12345678, 0x1234, 0x1234, { 0x12, 0x34, 0x56, 0x78, 0x90, 0xAB, 0xCD, 0xEF } };

// Supported file formats
const std::vector<std::wstring> SUPPORTED_FORMATS = {
    L".docx",
    L".docm",
    L".dotx",
    L".dotm"
};

// Menu item IDs
enum MenuItemID {
    MENU_CONVERT = 0,
    MENU_EXPORT = 1,
    MENU_OTHER = 2
};

// Resource IDs for localized strings
#define IDS_MENU_CONVERT 100
#define IDS_MENU_EXPORT 101
#define IDS_MENU_OTHER 102
#define IDS_SUBMENU_1 103
#define IDS_SUBMENU_2 104
#define IDS_SUBMENU_3 105

// Main shell extension class
class ContextMenuExtension : public IShellExtInit, public IContextMenu
{
public:
    ContextMenuExtension();
    virtual ~ContextMenuExtension();

    // IUnknown
    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObject);
    STDMETHODIMP_(ULONG) AddRef();
    STDMETHODIMP_(ULONG) Release();

    // IShellExtInit
    STDMETHODIMP Initialize(LPCITEMIDLIST pidlFolder, LPDATAOBJECT pDataObj, HKEY hProgID);

    // IContextMenu
    STDMETHODIMP GetCommandString(UINT_PTR idCmd, UINT uFlags, UINT* pwReserved, LPSTR pszName, UINT cchMax);
    STDMETHODIMP InvokeCommand(LPCMINVOKECOMMANDINFO pici);
    STDMETHODIMP QueryContextMenu(HMENU hMenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags);

private:
    LONG m_cRef;
    std::vector<std::wstring> m_selectedFiles;
    bool m_isDocxFile;
    
    // Helper methods
    bool IsSupportedFormat(const std::wstring& filePath);
    bool CheckNamespaceInDocx(const std::wstring& filePath, const std::wstring& targetNamespace);
    void ExecuteCommand(const std::wstring& command, const std::wstring& args);
    std::wstring GetLocalizedString(UINT resourceId);
};

// Windows 11 Explorer Command implementation
class ExplorerCommand : public IExplorerCommand
{
public:
    ExplorerCommand();
    virtual ~ExplorerCommand();

    // IUnknown
    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObject);
    STDMETHODIMP_(ULONG) AddRef();
    STDMETHODIMP_(ULONG) Release();

    // IExplorerCommand
    STDMETHODIMP GetTitle(IShellItemArray* psiItemArray, LPWSTR* ppszName);
    STDMETHODIMP GetIcon(IShellItemArray* psiItemArray, LPWSTR* ppszIcon);
    STDMETHODIMP GetToolTip(IShellItemArray* psiItemArray, LPWSTR* ppszInfotip);
    STDMETHODIMP GetCanonicalName(GUID* pguidCommandName);
    STDMETHODIMP GetState(IShellItemArray* psiItemArray, BOOL fOkToBeSlow, EXPCMDSTATE* pCmdState);
    STDMETHODIMP Invoke(IShellItemArray* psiItemArray, IBindCtx* pbc);
    STDMETHODIMP GetFlags(EXPCMDFLAGS* pFlags);
    STDMETHODIMP EnumSubCommands(IEnumExplorerCommand** ppEnum);

private:
    LONG m_cRef;
    std::vector<std::wstring> m_selectedFiles;
    bool m_isDocxFile;
    
    // Helper methods
    bool IsSupportedFormat(const std::wstring& filePath);
    bool CheckNamespaceInDocx(const std::wstring& filePath, const std::wstring& targetNamespace);
    void ExecuteCommand(const std::wstring& command, const std::wstring& args);
    std::wstring GetLocalizedString(UINT resourceId);
}; 