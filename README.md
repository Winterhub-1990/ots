# Windows Context Menu Extension

This project implements a Windows shell extension that adds custom context menu items for specific file formats (DOCX, DOCM, DOTX, DOTM). The extension works on both Windows 10 and Windows 11, with support for both 32-bit and 64-bit systems.

## Features

- Multi-culture menu text support (defaults to English)
- Windows 11 IExplorerCommand implementation
- Windows 10 legacy context menu support
- Dynamic menu item visibility based on DOCX file content
- Icon support for menu items
- Three submenu items with different shell commands

## Building the Project

1. Open the solution in Visual Studio 2022
2. Select the desired configuration (Debug/Release) and platform (x86/x64)
3. Build the solution (F7 or Build > Build Solution)

## Deployment

### Testing

1. Build the solution in Release mode
2. Open an elevated command prompt (Run as Administrator)
3. Navigate to the output directory (bin\x64\Release or bin\x86\Release)
4. Register the DLL:
   ```
   regsvr32 ContextMenuExtension.dll
   ```
5. To unregister:
   ```
   regsvr32 /u ContextMenuExtension.dll
   ```

### Production Deployment

For production deployment, you should create an installer that:
1. Copies the DLL to the appropriate system directory
2. Registers the COM server
3. Sets up the necessary registry entries

#### Required Registry Entries

The installer needs to create the following registry entries:

1. COM Server Registration:
   ```
   HKEY_CLASSES_ROOT\CLSID\{YOUR-GUID-HERE}
   HKEY_CLASSES_ROOT\CLSID\{YOUR-GUID-HERE}\InProcServer32
   ```

2. File Type Associations:
   ```
   HKEY_CLASSES_ROOT\.docx\shellex\ContextMenuHandlers\ContextMenuExtension
   HKEY_CLASSES_ROOT\.docm\shellex\ContextMenuHandlers\ContextMenuExtension
   HKEY_CLASSES_ROOT\.dotx\shellex\ContextMenuHandlers\ContextMenuExtension
   HKEY_CLASSES_ROOT\.dotm\shellex\ContextMenuHandlers\ContextMenuExtension
   ```

## Customization

### Adding/Changing File Formats

To add or modify supported file formats, edit the `SUPPORTED_FORMATS` array in `ContextMenuExtension.h`:

```cpp
const std::vector<std::wstring> SUPPORTED_FORMATS = {
    L".docx",
    L".docm",
    L".dotx",
    L".dotm"
};
```

### Localization

To add or modify localized strings:

1. Open `ContextMenuExtension.rc`
2. Add or modify string resources in the STRINGTABLE section
3. Add corresponding resource IDs in `resource.h`
4. Update the `GetLocalizedString` method in `ContextMenuExtension.cpp` to handle the new strings

### Menu Items

To modify menu items:

1. Update the menu item IDs in `resource.h`
2. Modify the `QueryContextMenu` method in `ContextMenuExtension.cpp`
3. Update the `InvokeCommand` method to handle the new menu items

## Troubleshooting

If the context menu items don't appear:

1. Verify the DLL is properly registered
2. Check the registry entries
3. Ensure the file format is in the supported formats list
4. Check if the namespace check is working correctly

## Dependencies

- Visual Studio 2022
- Windows SDK
- libzip (for DOCX file parsing)

## License

[Your License Here] 