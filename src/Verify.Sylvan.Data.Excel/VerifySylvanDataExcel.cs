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

        var sheets = Convert(reader, options).ToList();
        if (sheets.Count == 1)
        {
            var (csv, _) = sheets[0];
            return new(null, [new("csv", csv, null)]);
        }

        var targets = new List<Target>(sheets.Count);
        foreach (var sheet in sheets)
        {
            targets.Add(new("csv", sheet.Csv, reader.WorksheetName));
        }

        return new(null, targets);
    }

    static IEnumerable<(StringBuilder Csv, string? Name)> Convert(ExcelDataReader reader, CsvDataWriterOptions? options)
    {
        using var writer = new StringWriter();
        using var csvWriter = CsvDataWriter.Create(writer, options);
        csvWriter.Write(reader);
        yield return (writer.GetStringBuilder(), reader.WorksheetName);
    }
}