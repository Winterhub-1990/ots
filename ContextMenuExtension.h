#pragma once

#include <windows.h>
#include <shlobj.h>
#include <shobjidl_core.h>
#include <string>
#include <vector>
#include <memory>
#include <map>

// Target namespace to check in Word documents
const wchar_t* TARGET_NAMESPACE = L"http://schemas.bickard.com/office/extensions/xinvo/binding";

// Supported file formats
extern const std::vector<std::wstring> SUPPORTED_FORMATS;

// Command IDs for menu items
enum CommandId {
    CMD_CONVERT = 0,
    CMD_EXPORT,
    CMD_COMMAND1,
    CMD_COMMAND2,
    CMD_COMMAND3
};

// Generate a new GUID using Visual Studio's Create GUID tool
static const CLSID CLSID_ContextMenuExtension = 
{ 0x1234abcd, 0x1234, 0x1234, { 0x12, 0x34, 0x56, 0x78, 0x9a, 0xbc, 0xde, 0xf0 } };

class __declspec(uuid("{1234ABCD-1234-1234-1234-56789ABCDEF0}")) ContextMenuExtension : 
    public IExplorerCommand,
    public IContextMenu3,
    public IShellExtInit
{
public:
    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv);
    IFACEMETHODIMP_(ULONG) AddRef();
    IFACEMETHODIMP_(ULONG) Release();

    // IShellExtInit
    IFACEMETHODIMP Initialize(PCIDLIST_ABSOLUTE pidlFolder, IDataObject* pdtobj, HKEY hkeyProgID);

    // IExplorerCommand
    IFACEMETHODIMP GetTitle(IShellItemArray* psiItemArray, LPWSTR* ppszName);
    IFACEMETHODIMP GetIcon(IShellItemArray* psiItemArray, LPWSTR* ppszIcon);
    IFACEMETHODIMP GetToolTip(IShellItemArray* psiItemArray, LPWSTR* ppszInfotip);
    IFACEMETHODIMP GetCanonicalName(GUID* pguidCommandName);
    IFACEMETHODIMP GetState(IShellItemArray* psiItemArray, BOOL fOkToBeSlow, EXPCMDSTATE* pCmdState);
    IFACEMETHODIMP Invoke(IShellItemArray* psiItemArray, IBindCtx* pbc);
    IFACEMETHODIMP GetFlags(EXPCMDFLAGS* pFlags);
    IFACEMETHODIMP EnumSubCommands(IEnumExplorerCommand** ppEnum);

    // IContextMenu
    IFACEMETHODIMP QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags);
    IFACEMETHODIMP InvokeCommand(LPCMINVOKECOMMANDINFO pici);
    IFACEMETHODIMP GetCommandString(UINT_PTR idCmd, UINT uFlags, UINT* pwReserved, LPSTR pszName, UINT cchMax);

    // IContextMenu2
    IFACEMETHODIMP HandleMenuMsg(UINT uMsg, WPARAM wParam, LPARAM lParam);

    // IContextMenu3
    IFACEMETHODIMP HandleMenuMsg2(UINT uMsg, WPARAM wParam, LPARAM lParam, LRESULT* plResult);

    ContextMenuExtension();

protected:
    ~ContextMenuExtension();

private:
    long m_cRef;
    std::wstring m_selectedFile;
    bool m_hasTargetNamespace;

    bool CheckFileExtension(const std::wstring& filePath);
    bool CheckTargetNamespace(const std::wstring& filePath);
    std::wstring GetLocalizedString(UINT resourceId);
    void ExecuteCommand(CommandId cmdId, const std::wstring& filePath);
};

// Factory class
class ContextMenuExtensionFactory : public IClassFactory {
public:
    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv);
    IFACEMETHODIMP_(ULONG) AddRef();
    IFACEMETHODIMP_(ULONG) Release();

    // IClassFactory
    IFACEMETHODIMP CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppv);
    IFACEMETHODIMP LockServer(BOOL fLock);

    ContextMenuExtensionFactory();

protected:
    ~ContextMenuExtensionFactory();

private:
    long m_cRef;
};

// DLL exports
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppv);
STDAPI DllCanUnloadNow();
STDAPI DllRegisterServer();
STDAPI DllUnregisterServer(); 