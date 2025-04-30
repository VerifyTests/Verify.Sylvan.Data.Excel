public static class VerifySylvanDataExcelSettings
{
    public static void ExcelDataReaderOptions(this VerifySettings settings, ExcelDataReaderOptions options) =>
        settings.Context["VerifySylvanDataExcelDataReaderOptions"] = options;

    public static SettingsTask ExcelDataReaderOptions(this SettingsTask settings, ExcelDataReaderOptions options)
    {
        settings.CurrentSettings.ExcelDataReaderOptions(options);
        return settings;
    }

    internal static ExcelDataReaderOptions? GetExcelDataReaderOptions(this IReadOnlyDictionary<string, object> settings)
    {
        if (settings.TryGetValue("VerifySylvanDataExcelDataReaderOptions", out var value))
        {
            return (ExcelDataReaderOptions) value;
        }

        return null;
    }

    public static void CsvDataWriterOptions(this VerifySettings settings, CsvDataWriterOptions options) =>
        settings.Context["VerifySylvanDataCsvDataWriterOptions"] = options;

    public static SettingsTask CsvDataWriterOptions(this SettingsTask settings, CsvDataWriterOptions options)
    {
        settings.CurrentSettings.CsvDataWriterOptions(options);
        return settings;
    }

    internal static CsvDataWriterOptions? GetCsvDataWriterOptions(this IReadOnlyDictionary<string, object> settings)
    {
        if (settings.TryGetValue("VerifySylvanDataCsvDataWriterOptions", out var value))
        {
            return (CsvDataWriterOptions) value;
        }

        return null;
    }
}