# <img src="/src/icon.png" height="30px"> MsOfficeDiff

[![Build status](https://img.shields.io/appveyor/build/SimonCropp/MsOfficeDiff)](https://ci.appveyor.com/project/SimonCropp/MsOfficeDiff)
[![NuGet Status](https://img.shields.io/nuget/v/MsWordDiff.svg?label=MsWordDiff)](https://www.nuget.org/packages/MsWordDiff/)
[![NuGet Status](https://img.shields.io/nuget/v/MsExcelDiff.svg?label=MsExcelDiff)](https://www.nuget.org/packages/MsExcelDiff/)

A .NET tool that compares two Word documents side by side using Microsoft Word's built-in document comparison feature.<!-- singleLineInclude: intro. path: /src/intro.include.md -->

 * **MsWordDiff** (`diffword`) — Compare two Word documents using Microsoft Word's built-in document comparison feature
 * **MsExcelDiff** (`diffexcel`) — Compare two Excel workbooks using Microsoft's Spreadsheet Compare

**See [Milestones](../../milestones?state=closed) for release notes.**


## Open Source Maintenance Fee

This project participates in the [Open Source Maintenance Fee](https://opensourcemaintenancefee.org). The source code is freely available under the terms of the [license](license.txt). To support sustainable maintenance, use of the project's official binary releases in revenue-generating activities and all government agencies requires adherence to the [Open Source Maintenance Fee EULA](OsmfEula.txt). The fee is paid by [sponsoring Papyrine](https://github.com/sponsors/Papyrine).

This project uses [SponsorCheck](https://github.com/SimonCropp/SponsorCheck) to surface a build-time reminder in consuming projects that are not yet sponsoring.


## Requirements

 * Windows
 * .NET 10.0 or later


# MsWordDiff

Requires Microsoft Word installed (for MsWordDiff)


## Installation

```
dotnet tool install -g MsWordDiff
```

Nuget: https://nuget.org/packages/MsWordDiff/


## Usage

```
diffword <path1> <path2> [--quiet]
```

Where `<path1>` and `<path2>` are paths to the Word documents to compare.

Example:

```
diffword original.docx modified.docx
```

This will open Microsoft Word with a comparison view showing the differences between the two documents. The tool will wait until Word is closed before exiting.

<img src="/src/diff.png">


### Options


#### --quiet

Hide source documents in the comparison view, showing only the comparison document.

Example:

```
diffword original.docx modified.docx --quiet
```

<img src="/src/diff-quiet.png">


##### Showing Original and Revised Panes in Quiet Mode

When using `--quiet` mode, the original and revised document panes are initially hidden. They can be shown at any time within Word:

1. Click the **Review** tab in Word's ribbon
2. In the **Compare** group, click **Compare**
3. Select **Show Source Documents**
4. Choose the preferred view:
   - **Show Both** - Displays both original and revised documents in the right pane
   - **Show Original** - Shows only the original document
   - **Show Revised** - Shows only the revised document
   - **Hide Source Documents** - Returns to quiet mode (comparison only)

When both documents are visible, they scroll in sync with the comparison document for side-by-side review.

**Learn more:**
- [Compare and merge two versions of a document - Microsoft Support](https://support.microsoft.com/en-us/office/compare-and-merge-two-versions-of-a-document-f5059749-a797-4db7-a8fb-b3b27eb8b87e)
- [How to Compare Two Microsoft Word Documents](https://seekfast.org/blog/office-software/how-to-compare-two-microsoft-word-documents/)


### On first execution

On first execution the user will be prompted to choose their preferred UX mode (Standard or Quiet).

<img src="/src/firstRun.png">


### Configuration

The default behavior of options can be configured using settings commands.


#### View settings

```
diffword settings
```

This displays the settings file path and current settings.

```
C:\Users\SimonCropp\.config\MsWordDiff\settings.json
{
  "Quiet": true
}
```

By default, settings are stored in:

```
%USERPROFILE%\.config\MsWordDiff\settings.json
```


#### Configure default Quiet mode

```
diffword set-quiet <true|false>
```

Set the default value for the Quiet option. When set to `true`, the source documents will be hidden by default.

Examples:

```
diffword set-quiet true
diffword set-quiet false
```

Note: Command-line options always override settings file values.


## DiffEngine Integration

MsWordDiff is supported by [DiffEngine](https://github.com/VerifyTests/DiffEngine), making it available as a diff tool for [Verify](https://github.com/VerifyTests/Verify) snapshot testing.

When using Verify to test Word document output, DiffEngine can launch MsWordDiff to show differences between expected and actual documents.

To prioritize MsWordDiff as the diff tool:

```cs
DiffTools.UseOrder(DiffTool.MsWordDiff);
```

See [DiffEngine Diff Tool documentation](https://github.com/VerifyTests/DiffEngine/blob/main/docs/diff-tool.md#msworddiff) for more details.


## How It Works

The tool uses COM automation to:

 1. Open both documents in Microsoft Word (read-only)
 2. Generate a comparison document using Word's `CompareDocuments` feature
 3. Display the comparison with tracked changes highlighting differences
 4. Automatically close Word when the parent process exits


# MsExcelDiff

Requires Microsoft Office Professional Plus / Microsoft 365 Apps for Enterprise as [Spreadsheet Compare](https://support.microsoft.com/en-us/office/compare-two-versions-of-a-workbook-by-using-spreadsheet-compare-0e1627fd-ce14-4c33-9ab1-8ea82c6a5a7e) is only included in these editions


## Installation

```
dotnet tool install -g MsExcelDiff
```

Nuget: https://nuget.org/packages/MsExcelDiff/


## Usage

```
diffexcel <path1> <path2>
```

Where `<path1>` and `<path2>` are paths to the Excel workbooks to compare.

Example:

```
diffexcel original.xlsx modified.xlsx
```

This will open Microsoft Spreadsheet Compare showing the differences between the two workbooks. The tool will wait until Spreadsheet Compare is closed before exiting.

<img src="/src/excelDiff.png">


### Spreadsheet Compare Availability

Spreadsheet Compare is only included with certain Office editions:

 * Microsoft 365 Apps for Enterprise (formerly Office 365 ProPlus)
 * Office Professional Plus (2013, 2016, 2019, 2021, 2024)

It is **not** included with:

 * Microsoft 365 Personal or Family
 * Office Home & Student
 * Office Home & Business
 * Office Standard


### Configuration


#### View settings

```
diffexcel settings
```

This displays the settings file path and current settings.

By default, settings are stored in:

```
%USERPROFILE%\.config\MsExcelDiff\settings.json
```


#### Configure Spreadsheet Compare path

If Spreadsheet Compare is installed in a non-standard location, set the path:

```
diffexcel set-path "C:\Custom\Path\SPREADSHEETCOMPARE.EXE"
```

To clear the override and revert to auto-detection:

```
diffexcel set-path ""
```


## How It Works

The tool:

 1. Writes both file paths to a temp file (one per line)
 2. Launches `SPREADSHEETCOMPARE.EXE` with the temp file as an argument
 3. For Click-to-Run Office installs, launches via `AppVLP.exe` (the App-V virtualization layer)
 4. Creates a Windows Job Object to terminate Spreadsheet Compare when the parent process exits
 5. Waits until Spreadsheet Compare is closed before exiting
