Attribute VB_Name = "VBASourceControlMacros"
' =====================================================================================================================
' source-control-vba v0.1
' Copyright (c) 2020 Jonas R. Glesaaen (jonas@glesaaen.com)
'
' Helper macros to run the functions in VBASourceControl.
'
' Author: Jonas R. Glesaaen (jonas@glesaaen.com)
' License: MIT
' =====================================================================================================================

'@Folder("SourceControl")
'@ManualUpdate("True")
Option Explicit
Option Private Module

' Configuration
Private Const BackupBeforeImport As Boolean = True
Private Const ClearFolderBeforeExport As Boolean = True
Private Const UseSubfolders As Boolean = True

Private Const ExportImportNames As Boolean = True
Private Const CheckNamesOnly As Boolean = False

Private Const DebugPrinting As Boolean = True

'@Description("Export the project in this workbook.")
Public Sub ExportSourceCode()
Attribute ExportSourceCode.VB_Description = "Export the project in this workbook."

    If Not DebugPrinting Then
        VBASourceControl.DisableDebugPrinting
    End If

    VBASourceControl.Export _
        ThisWorkbook, _
        ClearContents:=ClearFolderBeforeExport, _
        WriteFolderStructure:=UseSubfolders, _
        ExportNames:=ExportImportNames
    
End Sub

'@Description("Import a project to this workbook.")
Public Sub ImportSourceCode()
Attribute ImportSourceCode.VB_Description = "Import a project to this workbook."

    If Not DebugPrinting Then
        VBASourceControl.DisableDebugPrinting
    End If

    VBASourceControl.Import _
        ThisWorkbook, _
        CreateBackup:=BackupBeforeImport, _
        Recursive:=UseSubfolders, _
        ImportNames:=ExportImportNames, _
        CheckNamesOnly:=CheckNamesOnly
    
End Sub
