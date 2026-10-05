# <img src="/src/icon.png" height="30px"> Verify.Sylvan.Data.Excel

[![Discussions](https://img.shields.io/badge/Verify-Discussions-yellow?svg=true&label=)](https://github.com/orgs/VerifyTests/discussions)
[![Build status](https://github.com/VerifyTests/Verify.Sylvan.Data.Excel/actions/workflows/build.yml/badge.svg)](https://github.com/VerifyTests/Verify.Sylvan.Data.Excel/actions/workflows/build.yml)
[![NuGet Status](https://img.shields.io/nuget/v/Verify.Sylvan.Data.Excel.svg)](https://www.nuget.org/packages/Verify.Sylvan.Data.Excel/)

Code provided by Cédric Luthi https://github.com/0xced

Extends [Verify](https://github.com/VerifyTests/Verify) to allow verification of Excel documents via [Sylvan.Data.Excel](https://github.com/MarkPflug/Sylvan.Data.Excel/).<!-- singleLineInclude: intro. path: /docs/intro.include.md -->

Converts Excel documents (xls, xlsb and xlsx) to csv for verification.

Verifying a workbook produces:

 * The workbook itself, as `.verified.xlsx`, `.verified.xlsb` or `.verified.xls`. An xlsx or an xlsb is a zip package, so it is passed through [DeterministicIoPackaging](https://github.com/SimonCropp/DeterministicIoPackaging), which removes the timestamps and the entry order that differ from one save to the next. This can be omitted with [`ExcludeTargets`](#choosing-what-is-verified).
 * An info file, `.verified.txt`, with the names of its sheets.
 * A csv of each sheet, named after the sheet: `#Sheet1.verified.csv`, `#Sheet2.verified.csv`, etc. A hidden sheet has a csv as any other, and is named under `HiddenSheets` in the info file. Options passed with `ExcelDataReaderOptions` are used as they are, so with those a hidden sheet is read only when `ReadHiddenWorksheets` is set. A reader that is passed in directly reads what it was created to read, and does not say which sheets are hidden. A sheet is a page, numbered by where it is among all the sheets of the workbook with hidden sheets counted, so `PagesToInclude` leaves out the csv of a sheet. A hidden sheet is counted whether or not the reader is asked to read it.

An `ExcelDataReader` passed to a verification has no workbook to hand, so only the info file and the csv files are produced for it.

The names of the csv files, and the setting that [leaves them out](#choosing-what-is-verified), are Verify's: the same for every plugin that splits a document into [derived targets](https://github.com/VerifyTests/Verify/blob/main/docs/converter.md#source-and-derived-targets).

**See [Milestones](../../milestones?state=closed) for release notes.**


## Sponsors


### Entity Framework Extensions<!-- include: sponsors. path: /docs/sponsors.include.md -->

[Entity Framework Extensions](https://entityframework-extensions.net/?utm_source=simoncropp&utm_medium=Verify.Sylvan.Data.Excel) is a major sponsor and is proud to contribute to the development this project.

[![Entity Framework Extensions](https://raw.githubusercontent.com/VerifyTests/Verify.Sylvan.Data.Excel/refs/heads/main/docs/zzz.png)](https://entityframework-extensions.net/?utm_source=simoncropp&utm_medium=Verify.Sylvan.Data.Excel)

### Developed using JetBrains IDEs

[![JetBrains logo.](https://raw.githubusercontent.com/VerifyTests/Verify.Sylvan.Data.Excel/main/docs/jetbrains.png)](https://jb.gg/OpenSourceSupport)<!-- endInclude -->


## NuGet

 * https://nuget.org/packages/Verify.Sylvan.Data.Excel


## Usage


### Enable Verify.Sylvan.Data.Excel

<!-- snippet: enable -->
<a id='snippet-enable'></a>
```cs
[ModuleInitializer]
public static void Initialize() =>
    VerifySylvanDataExcel.Initialize();
```
<sup><a href='/src/Tests/ModuleInitializer.cs#L3-L9' title='Snippet source file'>snippet source</a> | <a href='#snippet-enable' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Excel


#### Verify a file

<!-- snippet: VerifyExcel -->
<a id='snippet-VerifyExcel'></a>
```cs
[Test]
public Task VerifyExcel() =>
    VerifyFile("sample.xlsx");
```
<sup><a href='/src/Tests/Samples.cs#L7-L13' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyExcel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Verify a Stream

<!-- snippet: VerifyExcelStream -->
<a id='snippet-VerifyExcelStream'></a>
```cs
[Test]
public Task VerifyExcelStream()
{
    var stream = new MemoryStream(File.ReadAllBytes("sample.xlsx"));
    return Verify(stream, "xlsx");
}
```
<sup><a href='/src/Tests/Samples.cs#L122-L131' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyExcelStream' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Verify a ExcelDataReader

<!-- snippet: ExcelDataReader -->
<a id='snippet-ExcelDataReader'></a>
```cs
[Test]
public Task VerifyExcelDataReader()
{
    using var stream = File.OpenRead("sample.xlsx");
    using var reader = ExcelDataReader.Create(stream, ExcelWorkbookType.ExcelXml);
    return Verify(reader);
}
```
<sup><a href='/src/Tests/Samples.cs#L97-L107' title='Snippet source file'>snippet source</a> | <a href='#snippet-ExcelDataReader' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Example snapshot

The test `Samples.VerifyExcel` verifies a workbook with one sheet, named `Sheet1`. The names of its sheets are in `Samples.VerifyExcel.verified.txt`:

<!-- snippet: Samples.VerifyExcel.verified.txt -->
<a id='snippet-Samples.VerifyExcel.verified.txt'></a>
```txt
{
  SheetNames: [
    Sheet1
  ]
}
```
<sup><a href='/src/Tests/Samples.VerifyExcel.verified.txt#L1-L5' title='Snippet source file'>snippet source</a> | <a href='#snippet-Samples.VerifyExcel.verified.txt' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

And the sheet is in `Samples.VerifyExcel#Sheet1.verified.csv`:

<!-- snippet: Samples.VerifyExcel#Sheet1.verified.csv -->
<a id='snippet-Samples.VerifyExcel#Sheet1.verified.csv'></a>
```csv
0,First Name,Last Name,Gender,Country,Date,Age,Id,Formula
1,Dulce,Abril,Female,United States,2017-10-15,32,1562,1594
2,Mara,Hashimoto,Female,Great Britain,2016-08-16,25,1582,1607
3,Philip,Gent,Male,France,2015-05-21,36,2587,2623
4,Kathleen,Hanner,Female,United States,2017-10-15,25,3549,3574
5,Nereida,Magwood,Female,United States,2016-08-16,58,2468,2526
6,Gaston,Brumm,Male,United States,2015-05-21,24,2554,2578
```
<sup><a href='/src/Tests/Samples.VerifyExcel%23Sheet1.verified.csv#L1-L7' title='Snippet source file'>snippet source</a> | <a href='#snippet-Samples.VerifyExcel#Sheet1.verified.csv' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

The csv is named after its sheet even when the workbook has no other, so that a second sheet adds a file rather than renaming the first.


### CsvDataWriterOptions

Used to configure options for writing CSV data.

<!-- snippet: CsvDataWriterOptions -->
<a id='snippet-CsvDataWriterOptions'></a>
```cs
[Test]
public Task CsvDataWriterOptions()
{
    using var stream = File.OpenRead("sample.xlsx");
    var options = new CsvDataWriterOptions
    {
        Delimiter = '\t',
        Quote = '"',
    };

    return Verify(stream)
        .CsvDataWriterOptions(options);
}
```
<sup><a href='/src/Tests/Samples.cs#L142-L158' title='Snippet source file'>snippet source</a> | <a href='#snippet-CsvDataWriterOptions' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Choosing what is verified

The csv files are what was derived from the workbook, so Verify's [`ExcludeDerivedTargets`](https://github.com/VerifyTests/Verify/blob/main/docs/paged-documents.md#leaving-out-what-was-derived) leaves them out, and only the sheet names are verified. The sheets are then not read:

<!-- snippet: SheetNamesOnly -->
<a id='snippet-SheetNamesOnly'></a>
```cs
[Test]
public Task SheetNamesOnly() =>
    VerifyFile("sample.xlsx")
        .ExcludeDerivedTargets("csv");
```
<sup><a href='/src/Tests/Samples.cs#L133-L140' title='Snippet source file'>snippet source</a> | <a href='#snippet-SheetNamesOnly' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

`VerifierSettings.ExcludeDerivedTargets("csv")` does the same for every test.

`ExcludeTargets("xlsx")` leaves out the workbook instead, keeping the info file and the csv files. It is then not built.

A workbook has no pages, so the other settings for [paged documents](https://github.com/VerifyTests/Verify/blob/main/docs/paged-documents.md), `PageText` and `PagesToInclude`, have no effect on it.


## Reviewing changes

A change to a workbook is a change to several files: the workbook, the csv of each sheet that differs, and the info file when a sheet is added, removed or renamed. Verify tells the diff tool that the csv files and the info file were derived from the workbook, and [DiffEngineViewer](https://github.com/VerifyTests/DiffEngine/blob/main/docs/viewer.md#files-derived-from-a-document), which draws a workbook itself, shows them as one row and accepts them together. Other diff tools are given each file, as before.


## Migrating from 1.x

Version 2 moves to the [source and derived targets](https://github.com/VerifyTests/Verify/blob/main/docs/converter.md#source-and-derived-targets) of Verify 33.3. The API of this plugin is unchanged. What differs is that the workbook is now a snapshot too, and the name of a csv file, which Verify now decides, as it does for [every plugin that has moved](https://github.com/VerifyTests/Verify/blob/main/docs/paged-documents.md#migrating-from-33):

| | 1.x | 2.x |
| --- | --- | --- |
| The workbook | Not a snapshot | `Tests.Report.verified.xlsx` |
| A workbook with one sheet | `Tests.Report.verified.csv` | `Tests.Report#Sheet1.verified.csv` |
| A workbook with several sheets | `Tests.Report#Sheet1.verified.csv` | Unchanged |
| The info file | `Tests.Report.verified.txt` | Unchanged |

`VerifierSettings.ExcludeTargets("xlsx")` keeps the workbook out of the snapshots, as it was in 1.x.

A renamed snapshot shows as a new file and a pending delete. Accepting both, or running once with [AutoVerify](https://github.com/VerifyTests/Verify/blob/main/docs/autoverify.md), moves a test over. The content of the csv is unchanged, so source control shows it as a rename.

Two cases that are less common also differ:

 * A workbook that is itself a named target has its csv files named relative to it: one named `Attachment1` has `#Attachment1.Sheet1.verified.csv`, where 1.x did not use the name of the workbook. Passed to a verification as a `Target`, its info file takes the name as well: `#Attachment1.verified.txt`.
 * An `ExcelDataReader` that has been moved past its last sheet has no csv. In 1.x it had one, holding only a header row.
