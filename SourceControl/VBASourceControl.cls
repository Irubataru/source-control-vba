VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBASourceControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
' =====================================================================================================================
' source-control-vba v0.1
' Copyright (c) 2019-2020 Jonas R. Glesaaen (jonas@glesaaen.com)
'
' Utility functions for doing source control in VBA.
'
' Author: Jonas R. Glesaaen (jonas@glesaaen.com)
' License: MIT
' =====================================================================================================================

'@Folder("SourceControl")
'@ManualUpdate("True")
Option Explicit

Private Const ErrorCode As Long = 40725
Private Const NamesFilename As String = "names.csv"
Private Const QueriesFolderName As String = "queries"

Private Type MVBASourceControl
    ShouldPrintDebugInfo As Boolean
    NewDocuments As Collection
    NameCache As Dictionary
    NameDifferences As Collection
End Type

Private this As MVBASourceControl

' ---------------------------------------------------------------------------------------------------------------------
' --- Public methods --------------------------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------------------------------------------

' Export
' ---------------------------------------------------------------------------------------------------------------------

'@Description("Export the the VBProject of a workbook to a directory chosen by a file dialogue.")
Public Sub Export( _
       ByVal Book As Workbook, _
       Optional ByVal ClearContents As Boolean = False, _
       Optional ByVal WriteFolderStructure As Boolean = False, _
       Optional ByVal ExportNames As Boolean = False, _
       Optional ByVal ExportQueries As Boolean = False)
Attribute Export.VB_Description = "Export the the VBProject of a workbook to a directory chosen by a file dialogue."

    ' Choose which folder to store the results in
    With Application.FileDialog(msoFileDialogFolderPicker)
    
        .Title = "Save project in ..."
        
        .AllowMultiSelect = False
        .InitialFileName = Book.Path & "\"
        
        If Not .Show Then
            Err.Raise ErrorCode, _
                Source:="VBASourceControl.Export", _
                Description:="No export folder selected"
        End If
        
        Dim Directory As String
        Directory = .SelectedItems(1)
    
    End With
    
    ExportToFolder _
        Book, _
        Directory, _
        ClearContents:=ClearContents, _
        WriteFolderStructure:=WriteFolderStructure, _
        ExportNames:=ExportNames, _
        ExportQueries:=ExportQueries

End Sub

'@Description("Export the the VBProject of a workbook to a directory.")
Public Sub ExportToFolder( _
       ByVal Book As Workbook, _
       ByVal Directory As String, _
       Optional ByVal ClearContents As Boolean = False, _
       Optional ByVal WriteFolderStructure As Boolean = False, _
       Optional ByVal ExportNames As Boolean = False, _
       Optional ByVal ExportQueries As Boolean = False)
Attribute ExportToFolder.VB_Description = "Export the the VBProject of a workbook to a directory."
    
    With New Scripting.FileSystemObject
    
        If Not .FolderExists(Directory) Then
            Err.Raise ErrorCode, _
                Source:="VBASourceControl.ExportToFile", _
                Description:="Export directory does not exist."
        End If
    
        ' Potentially clear the previous files
        If ClearContents Then ClearContentsIn Directory
    
        Dim Component As VBIDE.VBComponent
        For Each Component In Book.VBProject.VBComponents: Do
        
                Dim ModuleName As String
                Select Case Component.Type
            
                Case VBIDE.vbext_ComponentType.vbext_ct_ClassModule
                    ModuleName = Component.Name & ".cls"
                
                Case VBIDE.vbext_ComponentType.vbext_ct_StdModule
                    ModuleName = Component.Name & ".bas"
                
                Case VBIDE.vbext_ComponentType.vbext_ct_Document
                    ModuleName = Component.Name & ".doccls"
                    
                Case VBIDE.vbext_ComponentType.vbext_ct_MSForm
                    ModuleName = Component.Name & ".frm"
                    
                Case VBIDE.vbext_ComponentType.vbext_ct_ActiveXDesigner
                    Debug.Assert False
                
                Case Else
                    'TODO Treat remaining components
            
                End Select
            
                ' Skip if ModuleName is empty
                If ModuleName = VBA.vbNullString Then Exit Do
            
                DebugLog "Exporting " & ModuleName
            
                Dim SaveTo As String
                SaveTo = Directory
            
                ' If we want to write the Rubberduck folder structure we need to make
                ' sure that these subfolders exist
                If WriteFolderStructure Then
                    SaveTo = CreateComponentDirectory(Component, Directory)
                End If
            
                ExportComponent Component, .BuildPath(SaveTo, ModuleName)
        
            Loop While False: Next Component

    End With
    
    If ExportNames Then
        ExportNamesToFolder Book, Directory
    End If
    
    If ExportQueries Then
        ExportQueriesToFolder Book, Directory
    End If

    MsgBox "Project exported successfully", Buttons:=vbInformation, Title:="VBASourceControl"
       
End Sub

'@Description("Export workbook names to a folder.")
Public Sub ExportNamesToFolder(ByVal Book As Workbook, ByVal Directory As String)
Attribute ExportNamesToFolder.VB_Description = "Export workbook names to a folder."

    With New Scripting.FileSystemObject
    
        If Not .FolderExists(Directory) Then
            Err.Raise ErrorCode, _
                Source:="VBASourceControl.ExportToFile", _
                Description:="Export directory does not exist."
        End If
        
        If Book.Names.Count < 1 Then Exit Sub
        
        With .CreateTextFile(.BuildPath(Directory, NamesFilename))
        
            ' Format "{Name},{RefersTo},{Comment}"
            Dim Name As Excel.Name
            For Each Name In Book.Names
                .WriteLine Name.Name & "," & Name.RefersTo & "," & Name.Comment
            Next Name
        
        End With
    
    End With

End Sub

'@Description("Export workbook queries to a folder.")
Public Sub ExportQueriesToFolder(ByVal Book As Workbook, ByVal Directory As String)
Attribute ExportQueriesToFolder.VB_Description = "Export workbook queries to a folder."

    With New Scripting.FileSystemObject
    
        If Not .FolderExists(Directory) Then
            Err.Raise ErrorCode, _
                Source:="VBASourceControl.ExportToFile", _
                Description:="Export directory does not exist."
        End If
        
        Dim QueriesDirectory As String
        QueriesDirectory = .BuildPath(Directory, QueriesFolderName)
        
        ' First clear out the old folder if it exists
        If .FolderExists(QueriesDirectory) Then
    
            Dim Folder As Scripting.Folder
            Set Folder = .GetFolder(QueriesDirectory)
    
            Dim File As Scripting.File
            For Each File In Folder.Files
                If GetFileExtension(File.Path) = "query" Then
                    DebugLog "Deleting " & File.Name
                    .DeleteFile File.Path
                End If
            Next File
            
            If Folder.Files.Count + Folder.SubFolders.Count = 0 Then
                .DeleteFolder Folder.Path
            End If
            
        End If
    
        If Book.Queries.Count < 1 Then Exit Sub
        
        ' The export the queries
        If Not .FolderExists(QueriesDirectory) Then
            .CreateFolder QueriesDirectory
        End If
        
        Dim i As Long
        For i = 1 To Book.Queries.Count
            With .CreateTextFile(.BuildPath(QueriesDirectory, Book.Queries(i).Name & ".query"))
                DebugLog "Exporting " & Book.Queries(i).Name & ".query"
                .WriteLine Book.Queries(i).Formula
            End With
        Next i
    
    End With

End Sub

' Import
' ---------------------------------------------------------------------------------------------------------------------

'@Description("Import a VBA project from a directory chosen by a file dialogue and add it to Workbook's projects.")
Public Sub Import( _
       ByVal Book As Workbook, _
       Optional ByVal CreateBackup As Boolean = False, _
       Optional ByVal Recursive As Boolean = False, _
       Optional ByVal ImportNames As Boolean = False, _
       Optional ByVal CheckNamesOnly As Boolean = False)
Attribute Import.VB_Description = "Import a VBA project from a directory chosen by a file dialogue and add it to Workbook's projects."
    
    ' Choose which folder to import the results from
    With Application.FileDialog(msoFileDialogFolderPicker)
    
        .Title = "Import project from ..."
        
        .AllowMultiSelect = False
        .InitialFileName = Book.Path & "\"
        
        If Not .Show Then
            Err.Raise ErrorCode, _
                Source:="VBASourceControl.Import", _
                Description:="No import folder selected"
        End If
        
        Dim Directory As String
        Directory = .SelectedItems(1)
    
    End With
    
    ImportFromFolder _
        Book, _
        Directory, _
        CreateBackup:=CreateBackup, _
        Recursive:=Recursive, _
        ImportNames:=ImportNames, _
        CheckNamesOnly:=CheckNamesOnly
    
End Sub

'@Description("Import a VBA project from a directory and add it to Workbook's projects.")
Public Sub ImportFromFolder( _
       ByVal Book As Workbook, _
       ByVal Directory As String, _
       Optional ByVal CreateBackup As Boolean = False, _
       Optional ByVal Recursive As Boolean = False, _
       Optional ByVal ImportNames As Boolean = False, _
       Optional ByVal CheckNamesOnly As Boolean = False)
Attribute ImportFromFolder.VB_Description = "Import a VBA project from a directory and add it to Workbook's projects."
       
    With New Scripting.FileSystemObject
        If Not .FolderExists(Directory) Then
            Err.Raise ErrorCode, _
                Source:="VBASourceControl.ImportFromFile", _
                Description:="Import directory does not exist."
        End If
    End With
       
    ' Create a backup of the project if CreateBackup = True
    If CreateBackup Then
        If CreateBackup Then
            If Not Me.BackupWorkbook(Book) Then
                Err.Raise ErrorCode, _
                    Source:="VBASourceControl.Import", _
                    Description:="No backup file created."
            End If
        End If
    End If
    
    ' Delete all classes and modules in the project
    DeleteAllComponentsInProject Book, CreateBackup:=False, DeleteManuallyUpdated:=False
    Application.Wait VBA.Now + VBA.TimeValue("00:00:01")
    
    Set this.NewDocuments = New Collection
    ImportAllComponentsInFolder Book, Directory, Recursive:=Recursive
    
    Set this.NameDifferences = New Collection
    If ImportNames Then
        ImportNamesFromFolder Book, Directory, DryRun:=CheckNamesOnly
    End If
    
    ShowImportInformation
    
    Set this.NewDocuments = Nothing
    Set this.NameDifferences = Nothing
       
End Sub

'@Description("Import workbook names from a folder.")
Public Sub ImportNamesFromFolder( _
        ByVal Book As Workbook, _
        ByVal Directory As String, _
        Optional ByVal DryRun As Boolean = False)
Attribute ImportNamesFromFolder.VB_Description = "Import workbook names from a folder."

    With New Scripting.FileSystemObject
    
        If Not .FolderExists(Directory) Then
            Err.Raise ErrorCode, _
                Source:="VBASourceControl.ExportToFile", _
                Description:="Export directory does not exist."
        End If
        
        ' Exit if no file exists
        If Not .FileExists(.BuildPath(Directory, NamesFilename)) Then Exit Sub
        
        ' Cache all the workbook names
        CacheWorkbookNames Book
        
        With .OpenTextFile(.BuildPath(Directory, NamesFilename))
            Do Until .AtEndOfStream
                ImportWorkbookName Book, .ReadLine, DryRun:=DryRun
            Loop
        End With
        
        Set this.NameCache = Nothing
    
    End With

End Sub

' Code cleanup
' ---------------------------------------------------------------------------------------------------------------------

'@Description("Delete all VBA component in a specific folder")
Public Sub DeleteAllComponentsInFolder( _
        ByVal Book As Workbook, _
        ByVal Folder As String, _
        Optional ByVal CreateBackup As Boolean = False, _
        Optional ByVal DeleteManuallyUpdated As Boolean = False)
Attribute DeleteAllComponentsInFolder.VB_Description = "Delete all VBA component in a specific folder"

    If CreateBackup Then
        If Not Me.BackupWorkbook(Book) Then
            Err.Raise ErrorCode, _
                Source:="VBASourceControl.DeleteFolder", _
                Description:="No backup file created."
        End If
    End If

    Dim Component As VBIDE.VBComponent
    For Each Component In Book.VBProject.VBComponents: Do
    
        If VBA.Left$(Directory(Component), VBA.Len(Folder)) <> Folder Then Exit Do
        If (Not DeleteManuallyUpdated) And IsManuallyUpdatedComponent(Component) Then Exit Do
        
        Select Case Component.Type
        
            Case vbext_ct_ClassModule, vbext_ct_StdModule
                Book.VBProject.VBComponents.Remove Component
                
            Case vbext_ct_Document
                ' If the component is a document we can only clear the contents, not delete it
                With Component.CodeModule
                    .DeleteLines 1, .CountOfLines
                End With
        
        End Select
    
    Loop While False: Next Component

End Sub

'@Description("Delete all VBA components in a workbook.")
Public Sub DeleteAllComponentsInProject( _
        ByVal Book As Workbook, _
        Optional ByVal CreateBackup As Boolean = False, _
        Optional ByVal DeleteManuallyUpdated As Boolean = False)
Attribute DeleteAllComponentsInProject.VB_Description = "Delete all VBA components in a workbook."

    Me.DeleteAllComponentsInFolder _
        Book, _
        VBA.vbNullString, _
        CreateBackup:=CreateBackup, _
        DeleteManuallyUpdated:=DeleteManuallyUpdated

End Sub

'@Description("Create a backup of a workbook.")
Public Function BackupWorkbook(ByVal Book As Workbook) As Boolean
Attribute BackupWorkbook.VB_Description = "Create a backup of a workbook."

    ' Choose filename of the backup
    With Application.FileDialog(msoFileDialogSaveAs)
    
        .Title = "Save backup file as..."
        
        Dim DefaultFilename As String
        With New Scripting.FileSystemObject
            DefaultFilename = .BuildPath(Book.Path, .GetBaseName(Book.Name) & ".bak")
        End With
        
        .AllowMultiSelect = False
        .InitialFileName = DefaultFilename
        .FilterIndex = 2  ' Default to .xlsm, TODO: default to workbook filetype
        
        ' Bring up the file picker
        If Not .Show Then Exit Function
        If .SelectedItems.Count < 1 Then Exit Function
        
        Dim TargetFilename As String
        TargetFilename = .SelectedItems(1)
    
    End With
    
    On Error GoTo FailedToSaveCopyError
    Book.SaveCopyAs TargetFilename
    BackupWorkbook = True

Finally:
    Exit Function

FailedToSaveCopyError:
    BackupWorkbook = False
    Resume Finally

End Function

' ---------------------------------------------------------------------------------------------------------------------
' --- Utility functions -----------------------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------------------------------------------

'@Description("Return the folder annotation for a VBA component.")
Public Function Directory(ByVal Component As VBIDE.VBComponent) As String
Attribute Directory.VB_Description = "Return the folder annotation for a VBA component."

    If Component.CodeModule.CountOfLines = 0 Then Exit Function

    With New VBScript_RegExp_55.RegExp
        .Global = False
        .MultiLine = True
        .Pattern = "^'@Folder[( ]""(.*)""\)?$"
    
        Dim ComponentCode As String
        ComponentCode = Component.CodeModule.Lines(1, Component.CodeModule.CountOfLines)
        
        If Not .Test(ComponentCode) Then Exit Function
        
        Dim Matches As VBScript_RegExp_55.MatchCollection
        Set Matches = .Execute(ComponentCode)
    
    End With
    
    If Matches.Count < 1 Then Exit Function
    If Matches(0).SubMatches.Count < 1 Then Exit Function
    
    Directory = Matches(0).SubMatches(0)

End Function

' Configuration
' ---------------------------------------------------------------------------------------------------------------------

'@Description("Turn debug printing on.")
Public Sub EnableDebugPrinting()
Attribute EnableDebugPrinting.VB_Description = "Turn debug printing on."
    this.ShouldPrintDebugInfo = True
End Sub

'@Description("Turn debug printing off.")
Public Sub DisableDebugPrinting()
Attribute DisableDebugPrinting.VB_Description = "Turn debug printing off."
    this.ShouldPrintDebugInfo = False
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' --- Private methods -------------------------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------------------------------------------

' Debugging
' ---------------------------------------------------------------------------------------------------------------------

'@Description("Log message.")
Private Sub DebugLog(ByVal Message As String)
Attribute DebugLog.VB_Description = "Log message."
    
    If this.ShouldPrintDebugInfo Then
        Debug.Print "(VBASourceControl) " & Message
    End If

End Sub

'@Description("Show information about the import task, some is printed to the Immediate window.")
Private Sub ShowImportInformation()
Attribute ShowImportInformation.VB_Description = "Show information about the import task, some is printed to the Immediate window."

    Dim RunInformation As String
    RunInformation = "Project imported successfully"
    
    If this.NewDocuments.Count > 0 Or this.NameDifferences.Count > 0 Then
        RunInformation = RunInformation & VBA.vbNewLine & VBA.vbNewLine & _
            "See the debug log for additional information"
    End If
    
    Dim i As Long
    If this.NewDocuments.Count > 0 Then
    
        Debug.Print "New sheets created:"

        For i = 1 To this.NewDocuments.Count
            Debug.Print " - " & this.NewDocuments(i)
        Next i
    
    End If
    
    If this.NameDifferences.Count > 0 Then
    
        Debug.Print "Names updated:"
            
        For i = 1 To this.NameDifferences.Count
            Debug.Print " - " & this.NameDifferences(i)
        Next i
    
    End If
    
    MsgBox RunInformation, Buttons:=vbInformation, Title:="VBASourceControl"

End Sub

' VBA project utilities
' ---------------------------------------------------------------------------------------------------------------------

'@Description("Export a VBA component to a file.")
Private Sub ExportComponent(ByVal Component As VBIDE.VBComponent, ByVal Path As String)
Attribute ExportComponent.VB_Description = "Export a VBA component to a file."

    Select Case Component.Type
    
        Case VBIDE.vbext_ComponentType.vbext_ct_ClassModule, _
             VBIDE.vbext_ComponentType.vbext_ct_StdModule, _
             VBIDE.vbext_ComponentType.vbext_ct_MSForm
            Component.Export Path
            
        Case VBIDE.vbext_ct_Document
            ' For documents we only write the content, not e.g. attributes
            ' This is consistent with how Rubberduck does things, and will simplify the reading process
            With New FileSystemObject
                With .CreateTextFile(Path)
                    .Write Component.CodeModule.Lines(1, Component.CodeModule.CountOfLines)
                    .Close
                End With
            End With
            
    End Select

End Sub

'@Description("Import a VBA component from a file.")
Private Sub ImportComponent(ByVal Book As Workbook, ByVal Path As String)
Attribute ImportComponent.VB_Description = "Import a VBA component from a file."

    Dim Filename As String
    With New Scripting.FileSystemObject
        Filename = .GetFileName(Path)
    End With
    
    If Not IsVBAComponentFiletype(Filename) Then Exit Sub  ' Skip files that aren't VBA files
    If IsManuallyUpdatedSourceFile(Path) Then Exit Sub     ' Skip files annotated as manually updated
    
    Select Case ComponentTypeFromFilename(Filename)
    
        Case VBIDE.vbext_ComponentType.vbext_ct_ClassModule, _
             VBIDE.vbext_ComponentType.vbext_ct_StdModule, _
             VBIDE.vbext_ComponentType.vbext_ct_MSForm
            Book.VBProject.VBComponents.Import Path
            
        Case VBIDE.vbext_ct_Document
            ' For documents we have to create the sheet in the workbook and paste the content
            ImportDocument Book, Path

    End Select

End Sub

'@Description("Imports document code from a file and stores it in a worksheet.")
Private Sub ImportDocument(ByVal Book As Workbook, ByVal Path As String)
Attribute ImportDocument.VB_Description = "Imports document code from a file and stores it in a worksheet."

    With New Scripting.FileSystemObject
        Dim Filename As String
        Filename = .GetFileName(Path)
        
        Dim Basename As String
        Basename = .GetBaseName(Path)
    End With
    
    Dim Component As VBIDE.VBComponent
    Set Component = ComponentByName(Book.VBProject, Filename)
    
    ' If no worksheet component was found we need to create one
    If Component Is Nothing Then
    
        Dim Sheet As Worksheet
        Set Sheet = Book.Worksheets.Add(After:=Book.Worksheets(Book.Worksheets.Count))
        this.NewDocuments.Add Basename
        
        Sheet.Name = Basename
        Set Component = Book.VBProject.VBComponents(Sheet.CodeName)
        
        ' Make sure variable names persist
        Component.Name = Basename
        
    End If
    
    With Component.CodeModule
        .DeleteLines 1, .CountOfLines   ' Delete the existing lines, just in case
        .AddFromFile Path               ' Import lines from file
    End With
    
End Sub

'@Description("Checks whether the component code has a ManualUpdate annotation.")
Private Function IsManuallyUpdatedComponent(ByVal Component As VBIDE.VBComponent) As Boolean
Attribute IsManuallyUpdatedComponent.VB_Description = "Checks whether the component code has a ManualUpdate annotation."
    
    With Component.CodeModule
        If .CountOfLines = 0 Then Exit Function
        IsManuallyUpdatedComponent = HasManualUpdateAnnotation(.Lines(1, .CountOfLines))
    End With

End Function

'@Description("Checks whether the code in the file has a ManualUpdate annotation.")
Private Function IsManuallyUpdatedSourceFile(ByVal Path As String) As Boolean
Attribute IsManuallyUpdatedSourceFile.VB_Description = "Checks whether the code in the file has a ManualUpdate annotation."

    With New Scripting.FileSystemObject
        With .OpenTextFile(Path, IOMode:=ForReading)
            IsManuallyUpdatedSourceFile = HasManualUpdateAnnotation(.ReadAll)
        End With
    End With

End Function

'@Description("Checks whether the source code contains a ManualUpdate annotation.")
Private Function HasManualUpdateAnnotation(ByVal SourceCode As String) As Boolean
Attribute HasManualUpdateAnnotation.VB_Description = "Checks whether the source code contains a ManualUpdate annotation."

    With New VBScript_RegExp_55.RegExp
        .Global = False
        .MultiLine = True
        .Pattern = "^'@ManualUpdate\(""True""\)"
        HasManualUpdateAnnotation = .Test(SourceCode)
    End With

End Function

'@Description("Searches fora component by name in a VBA project, returns nothing if none exists.")
Private Function ComponentByName(ByVal Project As VBIDE.VBProject, ByVal Name As String) As VBIDE.VBComponent
Attribute ComponentByName.VB_Description = "Searches fora component by name in a VBA project, returns nothing if none exists."

    ' Remove possible file extension
    If VBA.InStr(Name, ".") > 0 Then
        Name = VBA.Left$(Name, VBA.InStr(Name, ".") - 1)
    End If
    
    Dim Component As VBIDE.VBComponent
    For Each Component In Project.VBComponents
        If Component.Name = Name Then
            Set ComponentByName = Component
            Exit Function
        End If
    Next Component

End Function

'@Description("Checks the extension of a file to see if it is a VBA component.")
Private Function IsVBAComponentFiletype(ByVal Filename As String) As Boolean
Attribute IsVBAComponentFiletype.VB_Description = "Checks the extension of a file to see if it is a VBA component."
    
    Select Case GetFileExtension(Filename)
        Case "bas", "cls", "doccls", "frm", "frx"
            IsVBAComponentFiletype = True
        Case Else
            IsVBAComponentFiletype = False
    End Select

End Function

'@Description("Parses the type of VBA component a file is based on its file extension.")
Private Function ComponentTypeFromFilename(ByVal Filename As String) As VBIDE.vbext_ComponentType
Attribute ComponentTypeFromFilename.VB_Description = "Parses the type of VBA component a file is based on its file extension."

    Dim FileExtension As String
    If VBA.InStr(Filename, ".") >= 1 Then
        FileExtension = VBA.Right$(Filename, VBA.Len(Filename) - VBA.InStrRev(Filename, "."))
    End If
    
    Select Case FileExtension
        Case "bas"
            ComponentTypeFromFilename = VBIDE.vbext_ComponentType.vbext_ct_StdModule
        Case "cls"
            ComponentTypeFromFilename = VBIDE.vbext_ComponentType.vbext_ct_ClassModule
        Case "doccls"
            ComponentTypeFromFilename = VBIDE.vbext_ComponentType.vbext_ct_Document
        Case "frm"
            ComponentTypeFromFilename = VBIDE.vbext_ComponentType.vbext_ct_MSForm
    End Select

End Function

'@Description("File extension from a filename.")
Private Function GetFileExtension(ByVal Filename As String) As String
Attribute GetFileExtension.VB_Description = "File extension from a filename."

    With New Scripting.FileSystemObject
        GetFileExtension = .GetExtensionName(Filename)
    End With

End Function

' Name utilities
' ---------------------------------------------------------------------------------------------------------------------

'@Description("Import workbook names from a string, inform about differences.")
Private Sub ImportWorkbookName(ByVal Book As Workbook, ByVal Line As String, Optional ByVal DryRun As Boolean = False)
Attribute ImportWorkbookName.VB_Description = "Import workbook names from a string, inform about differences."

    If VBA.InStr(Line, ",") < 1 Then Exit Sub

    Dim SubStrings() As String
    SubStrings = VBA.Split(Line, ",")
    
    If UBound(SubStrings) - LBound(SubStrings) < 2 Then Exit Sub
    
    Dim Name As String: Name = SubStrings(0)
    Dim RefersTo As String: RefersTo = SubStrings(1)
    Dim Comment As String: Comment = SubStrings(2)
    
    If Not this.NameCache.Exists(Name) Then
        this.NameDifferences.Add "Adding new name """ & Name & """ (" & RefersTo & ")"
    ElseIf this.NameCache(Name) <> RefersTo Then
        this.NameDifferences.Add "Updating name """ & Name & """ (" & this.NameCache(Name) & " => " & RefersTo & ")"
    End If
    
    ' Do not actually do anything if we are in a dry run
    If DryRun Then Exit Sub
    
    Book.Names.Add(Name:=Name, RefersTo:=RefersTo).Comment = Comment
    
End Sub

'@Description("Create a cache of the names in a workbook.")
Private Sub CacheWorkbookNames(ByVal Book As Workbook)
Attribute CacheWorkbookNames.VB_Description = "Create a cache of the names in a workbook."

    Set this.NameCache = New Dictionary
    
    Dim Name As Excel.Name
    For Each Name In Book.Names
        this.NameCache.Add Name.Name, Name.RefersTo
    Next Name

End Sub

' Folder utilities
' ---------------------------------------------------------------------------------------------------------------------

'@Description("Create subfolders for a component if they do not already exist.")
Private Function CreateComponentDirectory(ByVal Component As VBIDE.VBComponent, ByVal BasePath As String) As String
Attribute CreateComponentDirectory.VB_Description = "Create subfolders for a component if they do not already exist."

    Dim ComponentDirectory As String
    ComponentDirectory = Me.Directory(Component) & "." ' Final "." is a sentinel for the loop
    
    ' Nothing to do if we are writing to root
    If ComponentDirectory = "." Then
        CreateComponentDirectory = BasePath
        Exit Function
    End If
    
    With New Scripting.FileSystemObject
    
        ' Keep appending subfolders to path
        Dim Path As String: Path = BasePath
        
        ' Loop through every subfolder (separated by ".")
        Do While VBA.InStr(ComponentDirectory, ".") > 0
        
            ' Append current subfolder to the path
            Path = Path & "\" & VBA.Left$(ComponentDirectory, VBA.InStr(ComponentDirectory, ".") - 1)
            
            ' Create the folder if it doesn't already exist
            If Not .FolderExists(Path) Then
                .CreateFolder Path
            End If
            
            ' Strip this folder from the beginning of the directory tree
            ComponentDirectory = _
                VBA.Right$( _
                    ComponentDirectory, _
                    VBA.Len(ComponentDirectory) - VBA.InStr(ComponentDirectory, "."))
        Loop
    
    End With
    
    CreateComponentDirectory = Path

End Function

'@Description("Delete everything in Directory related to a VBA project excluding dot files and folders.")
Private Sub ClearContentsIn(ByVal Directory As String)
Attribute ClearContentsIn.VB_Description = "Delete everything in Directory related to a VBA project excluding dot files and folders."

    With New Scripting.FileSystemObject
    
        Dim File As Scripting.File
        For Each File In .GetFolder(Directory).Files
        
            ' Ignore dotfiles, and files that are not VBA components
            If (Not VBA.Left$(File.Name, 1) = ".") And (IsVBAComponentFiletype(File.Name) Or File.Name = NamesFilename) Then
                DebugLog "Deleting " & File.Name
                .DeleteFile File.Path
            End If
            
        Next File
    
        Dim Folder As Scripting.Folder
        For Each Folder In .GetFolder(Directory).SubFolders

            If Not VBA.Left$(Folder.Name, 1) = "." Then
                ClearContentsIn Folder.Path
                
                ' Only delete a folder if it is empty
                If Folder.Files.Count + Folder.SubFolders.Count = 0 Then
                    DebugLog "Deleting " & Folder.Name
                    .DeleteFolder Folder.Path
                End If
            End If
            
        Next Folder
    
    End With

End Sub

'@Description("Import all VBA components in the folder to a VBProject.")
Private Sub ImportAllComponentsInFolder( _
        ByVal Book As Workbook, _
        ByVal Directory As String, _
        Optional ByVal Recursive As Boolean = False)
Attribute ImportAllComponentsInFolder.VB_Description = "Import all VBA components in the folder to a VBProject."
        
    With New Scripting.FileSystemObject
    
        Dim File As Scripting.File
        For Each File In .GetFolder(Directory).Files
            ImportComponent Book, File.Path
        Next File
        
        If Not Recursive Then Exit Sub
        
        Dim Folder As Scripting.Folder
        For Each Folder In .GetFolder(Directory).SubFolders
            
            ' Ignore dotfile directories
            If VBA.Left$(Folder.Name, 1) <> "." Then
                ImportAllComponentsInFolder Book, Folder.Path, Recursive
            End If
            
        Next Folder
        
    End With

End Sub

' ---------------------------------------------------------------------------------------------------------------------
' --- Constructors / destructors --------------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------------------------------------------

'@Description("Class constructor, disabled for everything except the default instance.")
Private Sub Class_Initialize()
Attribute Class_Initialize.VB_Description = "Class constructor, disabled for everything except the default instance."

    If Not Me Is VBASourceControl Then
        Err.Raise ErrorCode, _
            Source:="VBASourceControl.Initialize", _
            Description:="VBASourceControl is a singleton, do not create copies."
    End If
    
    this.ShouldPrintDebugInfo = True

End Sub

'@Description("Class destructor.")
Private Sub Class_Terminate()
Attribute Class_Terminate.VB_Description = "Class destructor."
    Set this.NewDocuments = Nothing
End Sub
