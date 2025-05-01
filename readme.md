# <img src="/src/icon.png" height="30px"> Verify.Sylvan.Data.Excel

[![Discussions](https://img.shields.io/badge/Verify-Discussions-yellow?svg=true&label=)](https://github.com/orgs/VerifyTests/discussions)
[![Build status](https://ci.appveyor.com/api/projects/status/q1eqcnbptyjl24hp?svg=true)](https://ci.appveyor.com/project/SimonCropp/verify-sylvan-data-excel)
[![NuGet Status](https://img.shields.io/nuget/v/Verify.Sylvan.Data.Excel.svg)](https://www.nuget.org/packages/Verify.Sylvan.Data.Excel/)

Code provided by Cédric Luthi https://github.com/0xced

Extends [Verify](https://github.com/VerifyTests/Verify) to allow verification of Excel documents via [Sylvan.Data.Excel](https://github.com/MarkPflug/Sylvan.Data.Excel/).

Converts documents xlsx to csv for verification.

**See [Milestones](../../milestones?state=closed) for release notes.**


## NuGet package

https://nuget.org/packages/Verify.Sylvan.Data.Excel/


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
<sup><a href='/src/Tests/Samples.cs#L8-L14' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyExcel' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/Tests/Samples.cs#L32-L41' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyExcelStream' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/Tests/Samples.cs#L20-L30' title='Snippet source file'>snippet source</a> | <a href='#snippet-ExcelDataReader' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Example snapshot

<!-- snippet: Samples.VerifyExcel.verified.csv -->
<a id='snippet-Samples.VerifyExcel.verified.csv'></a>
```csv
0,First Name,Last Name,Gender,Country,Date,Age,Id,Formula
1,Dulce,Abril,Female,United States,2017-10-15,32,1562,1594
2,Mara,Hashimoto,Female,Great Britain,2016-08-16,25,1582,1607
3,Philip,Gent,Male,France,2015-05-21,36,2587,2623
4,Kathleen,Hanner,Female,United States,2017-10-15,25,3549,3574
5,Nereida,Magwood,Female,United States,2016-08-16,58,2468,2526
6,Gaston,Brumm,Male,United States,2015-05-21,24,2554,2578
```
<sup><a href='/src/Tests/Samples.VerifyExcel.verified.csv#L1-L7' title='Snippet source file'>snippet source</a> | <a href='#snippet-Samples.VerifyExcel.verified.csv' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->
 

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
<sup><a href='/src/Tests/Samples.cs#L43-L59' title='Snippet source file'>snippet source</a> | <a href='#snippet-CsvDataWriterOptions' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->
