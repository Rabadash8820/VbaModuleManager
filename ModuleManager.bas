Attribute VB_Name = "ModuleManager"
Option Explicit
Option Private Module

#Const MANAGING_WORD = 0
#Const MANAGING_EXCEL = 0

Private Const MY_NAME = "ModuleManager"
Private Const ERR_SUPPORTED_APPS = MY_NAME & " currently only supports Microsoft Word and Excel."

Dim allComponents As VBComponents
Dim fileSys As New FileSystemObject
Dim alreadySaved As Boolean

Public Sub ImportModules(FromDirectory As String, Optional ShowMsgBox As Boolean = True)
    'Cache some references
    'If the given directory does not exist then show an error dialog and exit
    Dim fromPath As String: fromPath = FromDirectory
    If Not fileSys.FolderExists(fromPath) Then
        fromPath = getFilePath() & "\" & fromPath
        If Not fileSys.FolderExists(fromPath) Then
            MsgBox "Could not locate import directory:  " & FromDirectory
            Exit Sub
        End If
    End If
    Dim dir As Folder: Set dir = fileSys.GetFolder(fromPath)
                
    'Import all VB code files from the given directory if any)
    Dim f As File
    Dim imports As New Scripting.Dictionary 'Must be qualified to distinguish it from an MS Word Dictionary
    Dim numFiles As Integer: numFiles = 0
    For Each f In dir.Files
        Dim dotIndex As String: dotIndex = InStrRev(f.Name, ".")
        Dim ext As String: ext = UCase(Right(f.Name, Len(f.Name) - dotIndex))
        Dim correctType As Boolean: correctType = (ext = "BAS" Or ext = "CLS" Or ext = "FRM")
        Dim allowedName As Boolean: allowedName = Left(f.Name, InStrRev(f.Name, ".") - 1) <> MY_NAME
        If correctType And allowedName Then
            numFiles = numFiles + 1
            Dim replaced As Boolean: replaced = doImport(f)
            imports.Add f.Name, replaced
        End If
    Next f
    
    'Show a success message box, if requested
    If ShowMsgBox Then
        Dim i As Integer
        Dim msg As String: msg = numFiles & " modules imported:" & vbCrLf & vbCrLf
        For i = 0 To imports.Count - 1
            msg = msg & "    " & imports.Keys()(i) & IIf(imports.Items()(i), " (replaced)", " (new)") & vbCrLf
        Next i
        Dim result As VbMsgBoxResult: result = MsgBox(msg, vbOKOnly)
    End If
End Sub
Public Sub ExportModules(ToDirectory As String)
    'Cache some references
    'If the given directory does not exist then show an error dialog and exit
    Dim toPath As String: toPath = ToDirectory
    If Not fileSys.FolderExists(toPath) Then
        toPath = getFilePath() & "\" & toPath
        If Not fileSys.FolderExists(toPath) Then _
            fileSys.CreateFolder (toPath)
    End If
    Dim dir As Folder: Set dir = fileSys.GetFolder(toPath)
    
    'Export all modules from this file (except default MS Office modules)
    Dim vbc As VBComponent
    Dim allComponents As VBComponents: Set allComponents = getAllComponents()
    For Each vbc In allComponents
        Dim correctType As Boolean: correctType = (vbc.Type = vbext_ct_StdModule Or vbc.Type = vbext_ct_ClassModule Or vbc.Type = vbext_ct_MSForm)
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

    'Remove all modules from this file (except default MS Office modules obviously)
    Dim removals As New Collection
    Dim vbc As VBComponent
    Dim numModules As Integer: numModules = 0
    Dim allComponents As VBComponents: Set allComponents = getAllComponents()
    For Each vbc In allComponents
        Dim correctType As Boolean: correctType = (vbc.Type = vbext_ct_StdModule Or vbc.Type = vbext_ct_ClassModule Or vbc.Type = vbext_ct_MSForm)
        If correctType And vbc.Name <> MY_NAME Then
            numModules = numModules + 1
            removals.Add vbc.Name
            allComponents.Remove vbc
        End If
    Next vbc
    
    'Set the saved flag to prevent a save event loop
    'Save file again now that all modules have been removed
    alreadySaved = True
    Call saveFile
        
    'Show a success message box
    If ShowMsgBox Then
        Dim item As Variant
        Dim msg As String: msg = numModules & " modules successfully removed:" & vbCrLf & vbCrLf
        For Each item In removals
            msg = msg & "    " & item & vbCrLf
        Next item
        msg = msg & vbCrLf & "Don't forget to remove any empty lines after the Attribute lines in .frm files..." _
                  & vbCrLf & "ModuleManager will never be re-imported or exported.  You must do this manually if desired." _
                  & vbCrLf & "NEVER edit code in the VBE and a separate editor at the same time!"
        Dim result As VbMsgBoxResult: result = MsgBox(msg, vbOKOnly)
    End If
End Sub

Private Function getFilePath() As String
    #If MANAGING_WORD Then
        getFilePath = ThisDocument.path
    #ElseIf MANAGING_EXCEL Then
        getFilePath = ThisWorkbook.path
    #Else
        Call raiseUnsupportedAppError
    #End If
End Function

Private Function getAllComponents() As VBComponents
    #If MANAGING_WORD Then
        Set getAllComponents = ThisDocument.VBProject.VBComponents
    #ElseIf MANAGING_EXCEL Then
        Set getAllComponents = ThisWorkbook.VBProject.VBComponents
    #Else
        Call raiseUnsupportedAppError
    #End If
End Function

Private Sub saveFile()
    #If MANAGING_WORD Then
        ThisDocument.save
    #ElseIf MANAGING_EXCEL Then
        ThisWorkbook.save
    #Else
        Call raiseUnsupportedAppError
    #End If
End Sub

Private Sub raiseUnsupportedAppError()
    Err.Raise Number:=vbObjectError + 1, Description:=ERR_SUPPORTED_APPS
End Sub

Private Function doImport(ByRef codeFile As File) As Boolean
    'Determine whether a module with this name already exists
    Dim Name As String: Name = Left(codeFile.Name, Len(codeFile.Name) - 4)
    On Error Resume Next
    Dim allComponents As VBComponents: Set allComponents = getAllComponents()
    Dim m As VBComponent: Set m = allComponents.item(Name)
    If Err.Number <> 0 Then _
        Set m = Nothing
    On Error GoTo 0
        
    'If so, remove it
    Dim alreadyExists As Boolean: alreadyExists = Not (m Is Nothing)
    If alreadyExists Then _
        allComponents.Remove m
    
    'Then import the new module
    allComponents.Import (codeFile.path)
    doImport = alreadyExists
End Function

Private Function doExport(ByRef module As VBComponent, dirPath As String) As Boolean
    'Determine whether a file with this component's name already exists
    Dim ext As String
    Select Case module.Type
        Case vbext_ct_MSForm
            ext = "frm"
        Case vbext_ct_ClassModule
            ext = "cls"
        Case vbext_ct_StdModule
            ext = "bas"
    End Select
    Dim filePath As String: filePath = dirPath & "\" & module.Name & "." & ext
    Dim alreadyExists As Boolean: alreadyExists = fileSys.FileExists(filePath)
        
    'If so, remove it (even if its ReadOnly)
    If alreadyExists Then
        Dim f As File: Set f = fileSys.GetFile(filePath)
        If (f.Attributes And 1) Then _
            f.Attributes = f.Attributes - 1 'The bitmask for ReadOnly file attribute
        fileSys.DeleteFile (filePath)
    End If
    
    'Then export the module
    'Remove it also, so that the file stays small (and unchanged according to version control)
    module.Export (filePath)
    doExport = alreadyExists
End Function