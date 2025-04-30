public static class VerifySylvanDataExcel
{
    public static bool Initialized { get; private set; }

    public static void Initialize()
    {
        if (Initialized)
        {
            throw new("Already Initialized");
        }

        Initialized = true;

        VerifierSettings.RegisterStreamConverter("xls", (_, target, settings) => Convert(target, ExcelWorkbookType.Excel, settings));
        VerifierSettings.RegisterStreamConverter("xlsb", (_, target, settings) => Convert(target, ExcelWorkbookType.ExcelBinary, settings));
        VerifierSettings.RegisterStreamConverter("xlsx", (_, target, settings) => Convert(target, ExcelWorkbookType.ExcelXml, settings));
        VerifierSettings.RegisterFileConverter<ExcelDataReader>(Convert);
    }

    private static ConversionResult Convert(Stream stream, ExcelWorkbookType workbookType, IReadOnlyDictionary<string, object> settings)
    {
        var readerOptions = settings.GetExcelDataReaderOptions();
        using var excelReader = ExcelDataReader.Create(stream, workbookType, readerOptions);
        return Convert(excelReader, settings);
    }

    private static ConversionResult Convert(ExcelDataReader excelReader, IReadOnlyDictionary<string, object> settings)
    {
        var writerOptions = settings.GetCsvDataWriterOptions();

        var targets = new List<Target>();
        do
        {
            var writer = new StringWriter();
            using var csvWriter = CsvDataWriter.Create(writer, writerOptions);
            csvWriter.Write(excelReader);
            targets.Add(new("csv", writer.ToString(), excelReader.WorksheetName));
        } while (excelReader.NextResult());

        return new(null, targets);
    }
}