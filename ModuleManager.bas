Attribute VB_Name = "ModuleManager"
Option Explicit
Option Private Module

Private Const MY_NAME = "ModuleManager"

Dim allComponents As VBComponents
Dim fileSys As New FileSystemObject

Public Sub importMacros()
    'Cache some references
    Dim thisFolder As Folder
    Set allComponents = ThisWorkbook.VBProject.VBComponents
    Set thisFolder = fileSys.GetFolder(ThisWorkbook.Path)
                
    'Import all macros from this workbook's folder (if any)
    Dim imports As New Dictionary
    Dim numFiles As Integer, f As File, dotIndex As String, ext As String, correctType As Boolean, allowedName As Boolean, replaced As Boolean
    numFiles = 0
    For Each f In thisFolder.Files
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
    Dim msg As String, result As VbMsgBoxResult, i As Integer
    msg = numFiles & " modules successfully imported:" & vbCr & vbCr
    For i = 0 To imports.Count - 1
        msg = msg & "    " & imports.Keys()(i) & IIf(imports.Items()(i), " (replaced)", " (new)") & vbCr
    Next i
    result = MsgBox(msg, vbOKOnly)
End Sub
Public Sub exportMacros()
    'Cache some references
    Dim thisFolder As Folder
    Set allComponents = ThisWorkbook.VBProject.VBComponents
    Set thisFolder = fileSys.GetFolder(ThisWorkbook.Path)
    
    'Export all modules from this workbook (except sheet/workbook modules)
    Dim vbc As VBComponent, correctType As Boolean
    For Each vbc In allComponents
        correctType = (vbc.Type = vbext_ct_StdModule Or vbc.Type = vbext_ct_ClassModule Or vbc.Type = vbext_ct_MSForm)
        If correctType Then _
            Call doExport(vbc)
    Next vbc
End Sub
Public Sub removeMacros()
    'Cache some references
    Dim thisFolder As Folder
    Set allComponents = ThisWorkbook.VBProject.VBComponents
    
    'Remove all modules from this workbook (except sheet/workbook modules)
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
    
    'Show a success message box
    Dim msg As String, result As VbMsgBoxResult, item As Variant
    msg = numModules & " modules successfully removed:" & vbCr & vbCr
    For Each item In removals
        msg = msg & "    " & item & vbCr
    Next item
    msg = msg & vbCr & "NEVER edit code in the .bas files and the VBE at the same time!" & vbCr
    msg = msg & "Don't forget to remove any empty lines after the Attribute lines in .frm files..." & vbCr
    msg = msg & "Note that changes made to ModuleManager outside the VBE will be overwritten and never imported."
    result = MsgBox(msg, vbOKOnly)
End Sub

Private Function doImport(ByRef macroFile As File) As Boolean
    'Determine whether a module with this name already exists
    Dim Name As String, m As VBComponent
    Name = Left(macroFile.Name, Len(macroFile.Name) - 4)
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
    allComponents.Import (macroFile.Path)
    doImport = alreadyExists
End Function
Private Function doExport(ByRef module As VBComponent) As Boolean
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
    filePath = ThisWorkbook.Path & "\" & module.Name & "." & ext
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



