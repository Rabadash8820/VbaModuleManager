# WaveAnalysisScripts
All the old wave analysis scripts from my research with Dr. Jordan Renna ("forked" from MEA Cruncher)

## Module Management
Paste the following code into the ThisWorkbook module of any Excel workbook, and import ModuleManager.bas.  This will enable automatic module management, i.e. auto-import on workbook open, auto-export on workbook save, and auto-remove on workbook close.

    Option Explicit

    Dim alreadySaved As Boolean

    Private Sub Workbook_Open()
        Call importMacros
    End Sub
    Private Sub Workbook_AfterSave(ByVal Success As Boolean)
        If Success Then _
            Call exportMacros
    End Sub
    Private Sub Workbook_BeforeClose(Cancel As Boolean)
        'Prevent a save event loop
        If alreadySaved Then
            alreadySaved = False
            Exit Sub
        End If
        
        'Remove all modules and save (so that modules are never saved with this workbook)
        alreadySaved = True
        Call removeMacros
        ThisWorkbook.Save
    End Sub
