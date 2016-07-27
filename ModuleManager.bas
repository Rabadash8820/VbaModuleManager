Attribute VB_Name = "ModuleManager"
Option Explicit
Option Private Module

Private Const MY_NAME = "ModuleManager"

Dim allComponents As VBComponents
Dim fileSys As New FileSystemObject
Dim alreadySaved As Boolean

Public Sub ImportModules(ByVal fromDirectory As String, Optional ShowMsgBox As Boolean = True)
    'Cache some references
    'If the given directory does not exist then show an error dialog and exit
    Dim path As String, dir As Folder
    Set allComponents = ThisWorkbook.VBProject.VBComponents
    path = fromDirectory
    If Not fileSys.FolderExists(path) Then
        path = ThisWorkbook.path & "\" & path
        If Not fileSys.FolderExists(path) Then
            MsgBox "Could not locate import directory:  " & fromDirectory
            Exit Sub
        End If
    End If
    Set dir = fileSys.GetFolder(path)
                
    'Import all VB code files from the given directory if any)
    Dim imports As New Dictionary
    Dim numFiles As Integer, f As File, dotIndex As String, ext As String, correctType As Boolean, allowedName As Boolean, replaced As Boolean
    numFiles = 0
    For Each f In dir.Files
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
    
    'Show a success message box, if requested
    If ShowMsgBox Then
        Dim msg As String, result As VbMsgBoxResult, i As Integer
        msg = numFiles & " modules imported:" & vbCr & vbCr
        For i = 0 To imports.Count - 1
            msg = msg & "    " & imports.Keys()(i) & IIf(imports.Items()(i), " (replaced)", " (new)") & vbCr
        Next i
        result = MsgBox(msg, vbOKOnly)
    End If
End Sub
Public Sub ExportModules(ByVal toDirectory As String)
    'Cache some references
    'If the given directory does not exist then show an error dialog and exit
    Dim path As String, dir As Folder
    Set allComponents = ThisWorkbook.VBProject.VBComponents
    path = toDirectory
    If Not fileSys.FolderExists(path) Then
        path = ThisWorkbook.path & "\" & path
        If Not fileSys.FolderExists(path) Then _
            fileSys.CreateFolder (path)
    End If
    Set dir = fileSys.GetFolder(path)
    
    'Export all modules from this workbook (except sheet/workbook modules)
    Dim vbc As VBComponent, correctType As Boolean
    For Each vbc In allComponents
        correctType = (vbc.Type = vbext_ct_StdModule Or vbc.Type = vbext_ct_ClassModule Or vbc.Type = vbext_ct_MSForm)
        If correctType And vbc.Name <> MY_NAME Then _
            Call doExport(vbc, dir.path)
    Next vbc
End Sub
Public Sub RemoveModules(Optional ShowMsgBox As Boolean = True)
    'Check the saved flag to prevent a save event loop
    If alreadySaved Then
        alreadySaved = False
        Exit Sub
    End If
        
    'Cache some references
    Set allComponents = ThisWorkbook.VBProject.VBComponents
                        
    'Remove all modules from this workbook (except sheet/workbook modules obviously)
    Dim removals As New Collection
    Dim numModules As Integer, vbc As VBComponent, correctType As Boolean
    numModules = 0
    For Each vbc In allComponents
        correctType = (vbc.Type = vbext_ct_StdModule Or vbc.Type = vbext_ct_ClassModule Or vbc.Type = vbext_ct_MSForm)
        If correctType And vbc.Name <> MY_NAME Then
            numModules = numModules + 1
            removals.Add vbc.Name
            allComponents.Remove vbc
        End If
    Next vbc
    
    'Set the saved flag to prevent a save event loop
        'Save workbook again now that all modules have been removed
    alreadySaved = True
    ThisWorkbook.Save
        
    'Show a success message box
    If ShowMsgBox Then
        Dim msg As String, result As VbMsgBoxResult, item As Variant
        msg = numModules & " modules successfully removed:" & vbCr & vbCr
        For Each item In removals
            msg = msg & "    " & item & vbCr
        Next item
        msg = msg & vbCr & "Don't forget to remove any empty lines after the Attribute lines in .frm files..." _
                  & vbCr & "ModuleManager will never be re-imported or exported.  You must do this manually if desired." _
                  & vbCr & "NEVER edit code in the VBE and a separate editor at the same time!"
        result = MsgBox(msg, vbOKOnly)
    End If
End Sub

Private Function doImport(ByRef codeFile As File) As Boolean
    'Determine whether a module with this name already exists
    Dim Name As String, m As VBComponent
    Name = Left(codeFile.Name, Len(codeFile.Name) - 4)
    On Error Resume Next
    Set m = allComponents.item(Name)
    If Err.Number <> 0 Then _
        Set m = Nothing
    On Error GoTo 0
        
    'If so, remove it
    Dim alreadyExists As Boolean
    alreadyExists = Not (m Is Nothing)
    If alreadyExists Then _
        allComponents.Remove m
    
    'Then import the new module
    allComponents.Import (codeFile.path)
    doImport = alreadyExists
End Function
Private Function doExport(ByRef module As VBComponent, ByVal dirPath As String) As Boolean
    'Determine whether a file with this component's name already exists
    Dim ext As String, filePath As String, alreadyExists As Boolean
    Select Case module.Type
        Case vbext_ct_MSForm
            ext = "frm"
        Case vbext_ct_ClassModule
            ext = "cls"
        Case vbext_ct_StdModule
            ext = "bas"
    End Select
    filePath = dirPath & "\" & module.Name & "." & ext
    alreadyExists = fileSys.FileExists(filePath)
        
    'If so, remove it (even if its ReadOnly)
    If alreadyExists Then
        Dim f As File
        Set f = fileSys.GetFile(filePath)
        If (f.Attributes And 1) Then _
            f.Attributes = f.Attributes - 1 'The bitmask for ReadOnly file attribute
        fileSys.DeleteFile (filePath)
    End If
    
    'Then export the module
    'Remove it also, so that the workbook file stays small (and unchanged according to version control)
    module.Export (filePath)
    doExport = alreadyExists
End Function
