# Module Manager

## Why?

The Module Manager makes it much easier to work with complex VBA projects by handling the automatic importing and exporting of VBA Modules, Class Modules, and UserForms.  The Module Manager imports modules (\*.bas, \*.frm, \*.frx, and \*.cls files) from specified directories when the workbook is opened, re-exports them to specified directories when the workbook is saved, and removes them when the workbook is closed. This provides the following benefits:

* Code is stored as text, allowing macros to be edited in the IDE of your choice, like Visual Studio or Notepad++. All changes will be imported to the VB Editor (VBE) the next time you open the workbook!
* Storing code as text (rather than inside a binary workbook file) also facilitates change tracking with version control software like Git or SVN.
* Storing code in files separate from the main workbook eases collaboration on multiple macros within the same workbook, and allows for more atomic commits.
* By removing modules when the workbook is closed, the Module Manager prevents duplication of code between the text files and the workbook itself, and reduces the size of the workbook while it is sitting around on your hard drive. This is especially powerful in combination with Git [Large File Storage (LFS)](https://git-lfs.github.com/).

## Setup

ModuleManager works with Office 2007 and later (not tested in 2003 or earlier). Module Manager only supports the Word and Excel VB Editors, as the other Office applications do not have "This" modules (like "ThisWorkbook" in Excel) and so cannot listen for file open/save/close events. The following screenshots show the setup for Office 2016 on a Windows 10 machine.

To see an example of a file completely setup for Module Manager, you can download one of the files in the `test/` folder (`word.docm` for Word or `excel.xlsm` for Excel).

1. __Import the ModuleManager module__ file into your file (Excel workbook or Word document). You can either:
  
    1. Copy-paste the [ModuleManager.bas](ModuleManager.bas) source code into a new Module within the VB Editor (VBE). __Make sure that you also remove the `Attribute VB_NAME` line from the top of the file afterward!__
    2. Clone/download this repo and import the `ModuleManager.bas` file directly. Within the VBE, in the Project Explorer view, right click anywhere under the name of your project and select "Import file...", as shown in the screenshot below. Select the `ModuleManager.bas` file that you just downloaded and click "Open".

        ![Import ModuleManager module](screenshots/import-module-manager.png)

    __Note, module management does not apply to the ModuleManager itself, so it will always be present in your file and will never be re-imported, exported, or removed.__

2. __Add necessary references.__  Within the VBE, select "Tools > References".  In the dialog box, make sure that the following references are selected.  If any other references are already selected, then you should probably leave those too!
    * `Visual Basic For Applications`
    * `Microsoft <application> x.x Object Library` (e.g. `Microsoft Word 16.0 Object Library`)
    * `OLE Automation`
    * `Microsoft Office x.x Object Library`
    * `Microsoft Scripting Runtime` (note that this and many other references may not be available in Office for Mac)
    * `Microsoft Visual Basic for Applications Extensibility x.x`

    ![Select Tools > References](screenshots/references-menu.png)  

    ![Select references to add from the dialog](screenshots/references-dialog.png)

3. __Enable developer macro settings.__  On Windows, click "File > Options > Trust Center > Trust Center Settings...".  In the dialog box, select "Macro Settings", then check "Enable all macros" (or "Disable all macros except digitally signed macros" if you know what you're doing), __and__ "Trust access to the VBA project object model".  

    ![Open Trust Center from File > Options](screenshots/macro-security-trust-center.png)  

    ![Enable macros from Trust Center's Macro Settings](screenshots/macro-security-trust-center-settings.png)

4. __Modify the compiler constant__ for your application in the `ModuleManager` module. Open that module from the Project Explorer, find the `#Const` line for your Office application, then change the `0` to a `1`. For example, if you are using Excel, then change `#Const MANAGING_EXCEL = 0` to `#Const MANAGING_EXCEL = 1`.

5. __Paste the following code__ for your application into the "This" module of your project (e.g., `ThisDocument` in Word or `ThisWorkbook` in Excel, visible in the Project Explorer). This module will contain the code that actually handles the file Open, Save, and Close events; without it, ModuleManager would just take up space! The comments below provide further instructions for customizing the ModuleManager.

    __Make sure you remove or comment out these statements when you are ready to provide your file to end users.__ Otherwise, they might get confused by message boxes about import/export errors and strange text files appearing in the file's directory. Macros should always be present in the file for your end users, as they probably should not be viewing/editing your macros anyway.

### Word

```vbnet
Option Explicit

Private Sub Document_Open()
    'Provide the Path to a directory with VBA code files. Paths may be relative or absolute.
    'You can add additional ImportModules() statements to import from multiple locations.
    'The boolean argument specifies whether or not to show a Message Box on completion.
    'Remove or comment out these statements when you are ready to provide this document to end users,
    'so that they don't get confused by message boxes about import errors.
    Call ImportModules(ThisDocument.Name & ".modules", ShowMsgBox:=True)
End Sub

Private Sub Document_Close()
    'Provide the Path to a directory where you want to export modules. Paths may be relative or absolute.
    'You can add additional ExportModules() statements to export to multiple locations.
    'Remove or comment out these statements when you are ready to provide this document to end users,
    'so that they don't get confused by the appearance of a bunch of mysterious code files upon saving!
    Call ExportModules(ThisDocument.Name & ".modules")
    
    'The boolean argument specifies whether or not to show a Message Box on completion.
    'Remove or comment out this statement when you are ready to provide this document to end users,
    'so that modules are not removed, and functionality is not broken the next time the document is opened.
    Call RemoveModules(ShowMsgBox:=True)
End Sub
```

### Excel

```vbnet
Option Explicit

Private Sub Workbook_Open()
    'Provide the path to a directory with VBA code files. Paths may be relative or absolute.
    'You can add additional ImportModules() statements to import from multiple locations.
    'The boolean argument specifies whether or not to show a Message Box on completion.
    'Remove or comment out these statements when you are ready to provide this workbook to end users,
    'so that they don't get confused by message boxes about import errors.
    Call ImportModules(ThisWorkbook.Name & ".modules", ShowMsgBox:=True)
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean)
    'Provide the path to a directory where you want to export modules. Paths may be relative or absolute.
    'You can add additional ExportModules() statements to export to multiple locations.
    'Remove or comment out these statements when you are ready to provide this workbook to end users,
    'so that they don't get confused by the appearance of a bunch of mysterious code files upon saving!
    Call ExportModules(ThisWorkbook.Name & ".modules")
End Sub

Private Sub Workbook_BeforeClose(ByRef Cancel As Boolean)
    'The boolean argument specifies whether or not to show a Message Box on completion.
    'Remove or comment out this statement when you are ready to provide this workbook to end users,
    'so that modules are not removed, and functionality is not broken the next time the workbook is opened.
    Call RemoveModules(ShowMsgBox:=True)
End Sub
```

## Support

Issues and support questions may be posted on this repository's [Issues page](https://github.com/DanwareCreations/VbaModuleManager/issues).

## License

[MIT](./LICENSE.txt)
