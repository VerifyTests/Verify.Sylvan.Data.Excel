
using Sylvan.Data.Csv;
using Sylvan.Data.Excel;

public class Samples
{
    #region VerifyExcel

    [Test]
    public Task VerifyExcel() =>
        VerifyFile("sample.xlsx");

    #endregion

    [Test]
    public Task MultipleSheets() =>
        VerifyFile(ProjectFiles.sample_multiple_sheets_xlsx.Path);

    // A workbook that is itself a named target, as one split out of some other document is.
    // Verify names its info file after it, and the csv of each sheet relative to it
    [Test]
    public Task NamedWorkbook()
    {
        var stream = new MemoryStream(File.ReadAllBytes(ProjectFiles.sample_multiple_sheets_xlsx.Path));
        return Verify(new Target("xlsx", stream, "Attachment1"));
    }

    #region ExcelDataReader

    [Test]
    public Task VerifyExcelDataReader()
    {
        using var stream = File.OpenRead("sample.xlsx");
        using var reader = ExcelDataReader.Create(stream, ExcelWorkbookType.ExcelXml);
        return Verify(reader);
    }

    #endregion

    // Past its last sheet a reader is on no sheet, so there is no csv, only the sheet names
    [Test]
    public Task ReaderPastLastSheet()
    {
        using var stream = File.OpenRead("sample.xlsx");
        using var reader = ExcelDataReader.Create(stream, ExcelWorkbookType.ExcelXml);
        while (reader.NextResult())
        {
        }

        return Verify(reader);
    }

    #region VerifyExcelStream

    [Test]
    public Task VerifyExcelStream()
    {
        var stream = new MemoryStream(File.ReadAllBytes("sample.xlsx"));
        return Verify(stream, "xlsx");
    }

    #endregion

    #region SheetNamesOnly

    [Test]
    public Task SheetNamesOnly() =>
        VerifyFile("sample.xlsx")
            .ExcludeDerivedTargets("csv");

    #endregion

    #region CsvDataWriterOptions

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

    #endregion
}