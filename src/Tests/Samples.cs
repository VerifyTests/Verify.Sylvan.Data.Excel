
using Sylvan.Data.Excel;

[TestFixture]
public class Samples
{
    #region VerifyExcel

    [Test]
    public Task VerifyExcel() =>
        VerifyFile("sample.xlsx");

    #endregion

    #region ExcelDataReader

    [Test]
    public Task VerifyExcelDataReader()
    {
        var stream = File.OpenRead("sample.xlsx");
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
}