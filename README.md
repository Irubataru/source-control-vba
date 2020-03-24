source-control-vba
==================

Utility module to simplify having your VBA code under source control. It also
makes use of the [Rubberduck][rubberduck] directory structure if that is used,
but is in no way required.

 * [Features](#features)
 * [Installation](#installation)
 * [Dependencies](#dependencies)
 * [API](#api)
 * [Inspiration](#inspiration)

## Features

The module adds functionality for importing and exporting a VBA project. For the
export one can choose whether or not to create the same directory tree as the
one you have already specified for Rubberduck. When importing you can tell the
module whether or not it should crawl sub folders for files. See the API for
more information about the options these commands support.

## Installation

To install the module simply import the
[`VBASourceControl`](SourceControl/VBASourceControl.cls) file into your project.
You can optionally also import
[`VBASourceControlMacros`](SourceControl/VBASourceControlMacros.bas) if you want
some pre-made macros you can run to import and export.

## Dependencies

The module depends on the following VBA references (added through
`Tools->References`)

 * `Microsoft Scripting Runtime`
   * `FileSystemObject` for writing to and reading from files
   * `Dictionary` for storing workbook name cache
 * `Microsoft Visual Basic for Applications Extensibility 5.3`
   * `VBProject`, `VBComponent`, ... and all the other things necessary to work
     with VBA code projects
 * `Microsoft VBScript Regular Expressions 5.5`
   * `RegExp` for parsing Rubberduck folder structure and other annotations

## API

### The `VBASourceControl` class

This is a wrapper class for the following functions. It is a class so that I can
create a strict namespace, meaning that none of the functions leak out into your
project's namespace.

---

### Configuration

The class has a couple of constants which can be configured

 * `ErrorCode` Which error code VBA errors are reported as (default: `40725`)
 * `NamesFilename` The filename the workbook names are stored in (default:
   `names.csv`)
 * `QueriesFolderName` The folder to store the queries in (default: `queries`)

---

### `Export` function

```vb
Public Sub Export( _
       ByVal Book As Workbook, _
       Optional ByVal ClearContents As Boolean = False, _
       Optional ByVal WriteFolderStructure As Boolean = False, _
       Optional ByVal ExportNames As Boolean = False, _
       Optional ByVal ExportQueries As Boolean = False)
```

Exports the VBA project stored in `Book`. The method will bring up a file
dialogue window asking the user where to export the project to. The result of
that argument is passed on to the [`ExportToFolder`](#exporttofolder-function)
function, see that item for information about this function's arguments.

---

### `ExportToFolder` function

```vb
Public Sub ExportToFolder( _
       ByVal Book As Workbook, _
       ByVal Directory As String, _
       Optional ByVal ClearContents As Boolean = False, _
       Optional ByVal WriteFolderStructure As Boolean = False, _
       Optional ByVal ExportNames As Boolean = False, _
       Optional ByVal ExportQueries As Boolean = False)
```

Exports the VBA project stored in `Book` to a folder taken as an argument.

#### Arguments

```vb
ByVal Book As Workbook
```

The workbook the project that is exported is stored in.


```vb
ByVal Directory As String
```

The directory to export the project to. Throws an error if this directory
doesn't exist.

```vb
Optional ByVal ClearContents As Boolean = False
```

Whether or not to delete the contents of the folder you export to before
exporting. This makes it easier to e.g. track renaming of classes and modules as
you do not end up with duplicates in these situations. There are still things
that will not be deleted:

 * Files in the root directory that is not a VBA component (filetype does not
   end in `.bas`, `.cls`, or `.doccls`.
 * Files and folders whose name starts with `"."`, commonly referred to as
   dotfiles. This will ensure that e.g. your git directory will remain intact.

```vb
Optional ByVal WriteFolderStructure As Boolean = False
```

Whether or not to export the files including the Rubberduck folder structure. In
Rubberduck folders are annotated by having a comment `'@Folder("Foo.Bar")` at
the top of your file. This module reads this and will in this case place the
exported file in directory `[project-root]/Foo/Bar/`.

```vb
Optional ByVal ExportNames As Boolean = False
```

Whether or not to also export workbook names. See the
[`ExportNamesToFolder`](#exportnamestofolder-function) function for
documentation on the format. The directory passed to the export function is
`[project-root]`.

```vb
Optional ByVal ExportQueries As Boolean = False
```

Whether or not to also export the workbook queries. See the
[`ExportQueriesToFolder`](#exportqueriestofolder-function) function for
documentation. The directory passed to the export function is
`[project-root]/[queries-folder]`.

---

### `ExportNamesToFolder` function

```vb
Public Sub ExportNamesToFolder( _
        ByVal Book As Workbook, _
        ByVal Directory As String)
```

Exports names in a workbook to a CSV file. The format is

```csv
{Name},{RefersTo},{Comment}
```

Where each of these are a variable of the `Excel.Name` class. The file is saved
to `{Directory}/names.csv`, but this can be configured in the class headers.

#### Arguments

```vb
ByVal Book As Workbook
```

The workbook from which to export the names.

```vb
ByVal Directory As String
```

The directory to store the names file in.

---

### `ExportQueriesToFolder` function

```vb
Public Sub ExportQueriesToFolder( _
        ByVal Book As Workbook, _
        ByVal Directory As String)
```

Exports the queries in a workbook to separate files in the queries folder. The
queries are saved as text at<br/>
`{Directory}/[query-name].query`.

#### Arguments

```vb
ByVal Book As Workbook
```

The workbook from which to export the queries.

```vb
ByVal Directory As String
```

The directory to store the queries in.

---

### `Import` function

```vb
Public Sub Import( _
       ByVal Book As Workbook, _
       Optional ByVal CreateBackup As Boolean = False, _
       Optional ByVal Recursive As Boolean = False, _
       Optional ByVal ImportNames As Boolean = False, _
       Optional ByVal CheckNamesOnly As Boolean = False)
```

Imports VBA components from a directory to the VBA project of a workbook. The
directory from which to import the source is selected through a file dialogue.
The result of that selection is passed on to
[`ImportFromFolder`](#importfromfolder-function), see that function's
documentation for information about the function arguments.

---

### `ImportFromFolder` function

```vb
Public Sub ImportFromFolder( _
       ByVal Book As Workbook, _
       ByVal Directory As String, _
       Optional ByVal CreateBackup As Boolean = False, _
       Optional ByVal Recursive As Boolean = False, _
       Optional ByVal ImportNames As Boolean = False, _
       Optional ByVal CheckNamesOnly As Boolean = False)
```

Imports VBA components from a directory to the VBA project of a workbook. If a
VBA component has the annotation `'@ManualUpdate("True")` anywhere, this process
will be skipped for these components. The files in this module is tagged as such
because VBA has issues removing modules that are running.

Although the import deletes all classes and modules (not annotated as manually
updated), it does not delete document components and workbook names. It will
overwrite information about these if they exist in the directory, but will not
remove non-existing ones. For sheets the sheet will remain, but it will be
stripped of its code.

When the run is over, special information about sheets and names will be
available in the Immediate window of VBE. If no information appears that means
that these none of these types of changes occurred. This currently shows:

1. Imported sheets that didn't exist before.
2. Imported names that didn't exist before.
3. Changes to names after import.

I chose to always display name changes as they are slightly more fragile than
code changes in my opinion.

#### Arguments

```vb
ByVal Book As Workbook
```

The workbook to import the VBA project to.

```vb
ByVal Directory As String
```

The directory to import the project from. Throws an error if this directory
doesn't exist.

```vb
Optional ByVal CreateBackup As Boolean = False
```

Whether or not to create a backup of the workbook before the project is
imported. This backup will be a backup of the entire workbook, not just the VBA
code. If this option is true, and it fails to create a copy, the rest will not
run.

```vb
Optional ByVal Recursive As Boolean = False
```

Whether or not to (recursively) include files from sub folders. Similar to how
the folder cleaning works this will only include VBA components, and will skip
dotfiles.

```vb
Optional ByVal ImportNames As Boolean = False
```

Whether or not to also import workbook names. This will overwrite existing
names, but will not delete names not in the import file.

```vb
Optional ByVal CheckNamesOnly As Boolean = False
```

Whether or not do only do a "dry run" of the name import, meaning that no names
are actually imported, it only populates the list of changes shown after the
import. Has no effect if `ImportNames = False`.

---

### `ImportNamesFromFolder` function

```vb
Public Sub ImportNamesFromFolder( _
        ByVal Book As Workbook, _
        ByVal Directory As String, _
        Optional DryRun As Boolean = False)
```

Imports names to a workbook.

#### Arguments

```vb
ByVal Book As Workbook, _
```

The worksheet to import the names to.

```vb
ByVal Directory As String
```

The directory to import it finds the name file in. Assumes that the file
`{Directory}/names.csv` exists. The exact name can be changed in the class
configuration.

```vb
Optional DryRun As Boolean = False
```

Whether or not to do a dry run, meaning that information about name changes are
produced, but no names are actually imported. This currently only make sense
when called from [`ImportNamesFromFolder`](#importnamesfromfolder-function) as
this function doesn't print anything.

---

### `DeleteAllComponentsInFolder` function

```vb
Public Sub DeleteAllComponentsInFolder( _
        ByVal Book As Workbook, _
        ByVal Folder As String, _
        Optional ByVal CreateBackup As Boolean = False, _
        Optional ByVal DeleteManuallyUpdated As Boolean = False)
```

Loops through the VBA project in a workbook and deletes (or clears) every VBA
component that resides in a Rubberduck folder, including sub folders.

This function can be useful when e.g. writing a deployment module for your VBA
project, and want to e.g. delete everything in "Test" before passing the
workbook on.

```vb
ByVal Book As Workbook
```

The workbook to delete the components from.

```vb
ByVal Folder As String
```

The folder to delete components in. This is on the same format as the Rubberduck
folder annotation, i.e. `Foo.Bar`. If e.g. `Foo` is passed, and modules tagged
`Foo.Bar` will be deleted as they constitute a sub folder of `Foo`. If `Folder`
is an empty string, this will delete all components in the current project.

```vb
Optional ByVal CreateBackup As Boolean = False
```

Whether or not to create a backup of the file before deleting component. If the
code is unable to create a backup file it will throw an error, and the deletion
task is aborted.

```vb
Optional ByVal DeleteManuallyUpdated As Boolean = False
```

Whether or not to also delete components tagged with the
`'@ManualUpdate("True")` annotation.

---

### `DeleteAllComponentsInProject` function

```vb
Public Sub DeleteAllComponentsInProject( _
        ByVal Book As Workbook, _
        Optional ByVal CreateBackup As Boolean = False, _
        Optional ByVal DeleteManuallyUpdated As Boolean = False)
```

Deletes all components in the current project. This is just an alias for

```vb
DeleteAllComponentsInFolder(Book, VBA.vbNullString, CreateBackup, DeleteManuallyUpdated)
```

See that functions documentation for information about the arguments.

---

### `BackupWorkbook` function

```vb
Public Function BackupWorkbook(ByVal Book As Workbook) As Boolean
```

Creates a backup of a workbook. Will prompt the user with a file dialogue for
selecting where the backup will be saved.

#### Arguments

```vb
ByVal Book As Workbook
```

The workbook to create a backup of.

#### Return value

Returns whether or not it was successful in creating the backup file.

---

### `Directory` function

```vb
Public Function Directory(ByVal Component As VBIDE.VBComponent) As String
```

Returns the Rubberduck directory of a VBA component. The format is the same as
you specify, meaning that directories are separated by ".", and not "/" or "\".

#### Arguments

```vb
ByVal Component As VBIDE.VBComponent
```

The VBA component to read the folder annotation from.

#### Return value

The directory of the VBA component. If no directory has been specified, it
returns an empty string.

## Inspiration

Somewhat inspired by [vbaDeveloper][vbaDeveloper], but goes about solving some
of the same problems in a slightly different way:

#### Differences

 * This is just a VBA class, not an Excel add-in. It doesn't hook into e.g.
   saving and opening, and is intended to be much more manual. One can obviously
   make those hooks oneself if that is required.

 * It has some added functionality with respect to folders. Since
   [Rubberduck][rubberduck] gave us a way to have folders in VBE, I wanted to
   replicate this folder structure in the source code directory as well. Thus
   the import and export functions can both be passed a flag telling it whether
   to use folders or not.

## License

MIT

[vbaDeveloper]: https://github.com/hilkoc/vbaDeveloper
[rubberduck]: https://github.com/rubberduck-vba/Rubberduck
