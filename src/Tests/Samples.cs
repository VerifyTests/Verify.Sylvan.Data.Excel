
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

    // A hidden sheet is verified as any other, so it has a csv. What says it is hidden is
    // HiddenSheets in the info file.
    [Test]
    public Task HiddenSheet() =>
        Verify(FirstSheetHidden(), "xlsx");

    // Options that are passed in are used as they are, so a reader that is not asked for the
    // hidden sheets does not read them. They are still named in the info file.
    [Test]
    public Task HiddenSheetWithOptions() =>
        Verify(FirstSheetHidden(), "xlsx")
            .ExcelDataReaderOptions(new());

    // A sheet is a page, so PagesToInclude limits the csv files. The info file still names every
    // sheet, and the xlsx is still the whole workbook.
    [Test]
    public Task PagesToInclude() =>
        VerifyFile(ProjectFiles.sample_multiple_sheets_xlsx.Path)
            .PagesToInclude(_ => _ == 2);

    // A hidden sheet is counted as any other, so here it is the first page
    [Test]
    public Task HiddenSheetIsCountedByPagesToInclude() =>
        Verify(FirstSheetHidden(), "xlsx")
            .PagesToInclude(1);

    // It is counted when it is not read as well. The reader is not asked for hidden sheets here, so
    // the sheet it does read is still the second page.
    [Test]
    public Task HiddenSheetIsCountedWhenNotRead() =>
        Verify(FirstSheetHidden(), "xlsx")
            .ExcelDataReaderOptions(new())
            .PagesToInclude(_ => _ == 2);

    // A sheet that only code can unhide is verified as one Excel can
    [Test]
    public Task VeryHiddenSheet() =>
        Verify(FirstSheetHidden("veryHidden"), "xlsx");

    // The workbook with two sheets, with the first marked hidden in xl/workbook.xml
    static MemoryStream FirstSheetHidden(string state = "hidden")
    {
        var stream = new MemoryStream();
        using (var file = File.OpenRead(ProjectFiles.sample_multiple_sheets_xlsx.Path))
        {
            file.CopyTo(stream);
        }

        using (var archive = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry("xl/workbook.xml")!;
            System.Xml.Linq.XDocument xml;
            using (var read = entry.Open())
            {
                xml = System.Xml.Linq.XDocument.Load(read);
            }

            var sheet = xml.Descendants().First(_ => _.Name.LocalName == "sheet");
            sheet.SetAttributeValue("state", state);

            using var write = entry.Open();
            write.SetLength(0);
            xml.Save(write);
        }

        stream.Position = 0;
        return stream;
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