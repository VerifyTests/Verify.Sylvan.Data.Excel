
using Sylvan.Data.Csv;
using Sylvan.Data.Excel;

[TestFixture]
public class Samples
{
    #region VerifyExcel

    [Test]
    public Task VerifyExcel() =>
        VerifyFile("sample.xlsx");

    #endregion

    [Test]
    public Task MultipleSheets() =>
        VerifyFile("sample_multiple_sheets.xlsx");

    #region ExcelDataReader

    [Test]
    public Task VerifyExcelDataReader()
    {
        using var stream = File.OpenRead("sample.xlsx");
        using var reader = ExcelDataReader.Create(stream, ExcelWorkbookType.ExcelXml);
        return Verify(reader);
    }

    #endregion

    #region VerifyExcelStream

    [Test]
    public Task VerifyExcelStream()
    {
        var stream = new MemoryStream(File.ReadAllBytes("sample.xlsx"));
        return Verify(stream, "xlsx");
    }

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