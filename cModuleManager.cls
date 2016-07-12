VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cModuleManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'HIDDEN FIELDS
Private Const MY_NAME = "cModuleManager"
Private WithEvents mWkbk As Workbook
Private mShowImport As Boolean
Private mShowRemove As Boolean
Private mRelease As Boolean
Private mFolder As Folder
Private mComponents As VBComponents
Private mFileSys As New FileSystemObject
Private mSaved As Boolean

'EVENT HANDLERS
Private Sub mWkbk_BeforeSave(ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean)
    If mRelease Then Exit Sub
    Call exportModules
End Sub
Private Sub mWkbk_BeforeClose(ByRef Cancel As Boolean)
    If mRelease Then Exit Sub
        
    'Prevent a save event loop
    If mSaved Then
        mSaved = False
        Exit Sub
    End If

    'Remove all modules and save (so that modules are never saved with this workbook)
    mSaved = True
    Call removeModules
    mWkbk.Save
End Sub

'INTERFACE FUNCTIONS
Public Sub StartManaging(Optional FolderPath As String, Optional ShowImportMsgBox As Boolean = True, Optional ShowRemoveMsgBox As Boolean = True, Optional ReleaseMode As Boolean = False)
    'Intitialize private variables
    Set mWkbk = ThisWorkbook
    Set mComponents = mWkbk.VBProject.VBComponents
    mShowImport = ShowImportMsgBox
    mShowRemove = ShowRemoveMsgBox
    mRelease = ReleaseMode
    
    'Set the module folder
    Dim path As String
    path = mWkbk.path & "\" & FolderPath
    If Not mFileSys.FolderExists(path) Then _
        mFileSys.CreateFolder (path)
    Set mFolder = mFileSys.GetFolder(path)
    
    'Import any existing modules, if necessary
    If Not mRelease Then _
        Call importModules
End Sub

'HELPER FUNCTIONS
Private Sub importModules()
    'Import all macros from this workbook's folder (if any)
    Dim imports As New Dictionary
    Dim numFiles As Integer, f As File, dotIndex As String, ext As String, correctType As Boolean, allowedName As Boolean, replaced As Boolean
    numFiles = 0
    For Each f In mFolder.Files
        dotIndex = InStrRev(f.Name, ".")
        ext = UCase(Right(f.Name, Len(f.Name) - dotIndex))
        correctType = (ext = "BAS" Or ext = "CLS" Or ext = "FRM")
        allowedName = Left(f.Name, InStrRev(f.Name, ".") - 1) <> MY_NAME
        If correctType And allowedName Then
            numFiles = numFiles + 1
            replaced = doImport(f)
            imports.Add f.Name, replaced
        End If
    Next f
    
    'Show a success message box
    If mShowImport Then
        Dim msg As String, result As VbMsgBoxResult, i As Integer
        msg = numFiles & " modules successfully imported:" & vbCr & vbCr
        For i = 0 To imports.Count - 1
            msg = msg & "    " & imports.Keys()(i) & IIf(imports.Items()(i), " (replaced)", " (new)") & vbCr
        Next i
        result = MsgBox(msg, vbOKOnly)
    End If
End Sub
Private Sub exportModules()
    'Export all modules from this workbook (except sheet/workbook modules)
    Dim vbc As VBComponent, correctType As Boolean
    For Each vbc In mComponents
        correctType = (vbc.Type = vbext_ct_StdModule Or vbc.Type = vbext_ct_ClassModule Or vbc.Type = vbext_ct_MSForm)
        If correctType And vbc.Name <> MY_NAME Then _
            Call doExport(vbc)
    Next vbc
End Sub
Private Sub removeModules()
    'Remove all modules from this workbook (except sheet/workbook modules)
    Dim removals As New Collection
    Dim numModules As Integer, vbc As VBComponent, correctType As Boolean
    numModules = 0
    For Each vbc In mComponents
        correctType = (vbc.Type = vbext_ct_StdModule Or vbc.Type = vbext_ct_ClassModule Or vbc.Type = vbext_ct_MSForm)
        If correctType And vbc.Name <> MY_NAME Then
            numModules = numModules + 1
            removals.Add vbc.Name
            mComponents.Remove vbc
        End If
    Next vbc
    
    'Show a success message box
    If mShowRemove Then
        Dim msg As String, result As VbMsgBoxResult, item As Variant
        msg = numModules & " modules successfully removed:" & vbCr & vbCr
        For Each item In removals
            msg = msg & "    " & item & vbCr
        Next item
        msg = msg & vbCr & "Don't forget to remove empty lines after the Attribute lines in .frm files..." _
                  & vbCr & "NEVER edit code in the text files and the VBE at the same time!"
        result = MsgBox(msg, vbOKOnly)
    End If
End Sub

Private Function doImport(ByRef macroFile As File) As Boolean
    'Determine whether a module with this name already exists
    Dim Name As String, m As VBComponent
    Name = Left(macroFile.Name, Len(macroFile.Name) - 4)
    On Error Resume Next
    Set m = mComponents.item(Name)
    If Err.Number <> 0 Then _
        Set m = Nothing
    On Error GoTo 0
        
    'If so, remove it
    Dim alreadyExists As Boolean
    alreadyExists = Not (m Is Nothing)
    If alreadyExists Then _
        mComponents.Remove m
    
    'Then import the new module
    mComponents.Import (macroFile.path)
    doImport = alreadyExists
End Function
Private Function doExport(ByRef module As VBComponent) As Boolean
    'Determine whether a file with this component's name already exists
    Dim ext As String, filePath As String, exists As Boolean
    Select Case module.Type
        Case vbext_ct_MSForm
            ext = "frm"
        Case vbext_ct_ClassModule
            ext = "cls"
        Case vbext_ct_StdModule
            ext = "bas"
    End Select
    filePath = mFolder.path & "\" & module.Name & "." & ext
    exists = mFileSys.FileExists(filePath)

    'If so, remove it (even if its ReadOnly)
    If exists Then
        Dim f As File
        Set f = mFileSys.GetFile(filePath)
        If (f.Attributes And 1) Then _
            f.Attributes = f.Attributes - 1 'The bitmask for ReadOnly file attribute
        mFileSys.DeleteFile (filePath)
    End If
    
    'Then export the module
    'Remove it also, so that the workbook file stays small (and unchanged according to version control)
    module.Export (filePath)
    doExport = exists
End Function
