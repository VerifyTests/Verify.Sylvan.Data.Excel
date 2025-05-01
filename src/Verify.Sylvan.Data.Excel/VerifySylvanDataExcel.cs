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

    static ConversionResult Convert(Stream stream, ExcelWorkbookType type, IReadOnlyDictionary<string, object> settings)
    {
        var options = settings.GetExcelDataReaderOptions();
        using var reader = ExcelDataReader.Create(stream, type, options);
        return Convert(reader, settings);
    }

    static ConversionResult Convert(ExcelDataReader reader, IReadOnlyDictionary<string, object> settings)
    {
        var options = settings.GetCsvDataWriterOptions();

        var targets = new List<Target>();
        do
        {
            using var writer = new StringWriter();
            using var csvWriter = CsvDataWriter.Create(writer, options);
            csvWriter.Write(reader);
            targets.Add(new("csv", writer.GetStringBuilder(), reader.WorksheetName));
        } while (reader.NextResult());

        return new(null, targets);
    }
}