namespace VerifyTests;

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
        // Read from a copy, so the workbook is still whole to be the source once the reader is done
        // with it
        var bytes = new MemoryStream();
        stream.CopyTo(bytes);
        bytes.Position = 0;

        var options = settings.GetExcelDataReaderOptions();
        using var reader = ExcelDataReader.Create(bytes, type, options);
        return Convert(reader, settings, BuildSource(bytes, type, settings));
    }

    // The workbook itself, which is what ties the csv files to it for comparison and for the diff
    // tool. Not built when ExcludeTargets says it is not wanted.
    static Target? BuildSource(MemoryStream bytes, ExcelWorkbookType type, IReadOnlyDictionary<string, object> settings)
    {
        var extension = Extension(type);
        if (settings.IsTargetExcluded(extension))
        {
            return null;
        }

        var copy = new MemoryStream(bytes.ToArray());

        // xlsx and xlsb are zip packages, which carry timestamps and an entry order that differ
        // from one save to the next. xls is a compound file with neither, so it is kept as it is.
        if (type == ExcelWorkbookType.Excel)
        {
            return new(extension, copy);
        }

        return new(extension, DeterministicPackage.Convert(copy));
    }

    static string Extension(ExcelWorkbookType type)
    {
        if (type == ExcelWorkbookType.Excel)
        {
            return "xls";
        }

        if (type == ExcelWorkbookType.ExcelBinary)
        {
            return "xlsb";
        }

        return "xlsx";
    }

    // A reader passed in directly has no workbook to snapshot: the bytes it reads are not to hand
    static ConversionResult Convert(ExcelDataReader reader, IReadOnlyDictionary<string, object> settings) =>
        Convert(reader, settings, null);

    static ConversionResult Convert(ExcelDataReader reader, IReadOnlyDictionary<string, object> settings, Target? source)
    {
        var info = new Info
        {
            SheetNames = reader.WorksheetNames,
        };

        // Reading the sheets is the expensive part, so it is skipped when the csv files are
        // excluded, by ExcludeDerivedTargets or by ExcludeTargets
        List<Target> sheets = [];
        if (!settings.IsDerivedTargetExcluded("csv"))
        {
            var options = settings.GetCsvDataWriterOptions();
            sheets.AddRange(Convert(reader, options));
        }

        // Verify names the csv files relative to the target that was converted, and with a source
        // compares it first and tells the diff tool the csv files came from it
        return new(info, source, derived: sheets);
    }

    static IEnumerable<Target> Convert(ExcelDataReader reader, CsvDataWriterOptions? options)
    {
        // No name is no sheet: a workbook with none to read, or a reader moved past its last
        if (reader.WorksheetName is null)
        {
            yield break;
        }

        do
        {
            using var writer = new StringWriter();
            using var csvWriter = CsvDataWriter.Create(writer, options);
            csvWriter.Write(reader);
            // Named after its sheet even when it is the only one, so that a second sheet adds a
            // file rather than renaming the first
            yield return new("csv", writer.GetStringBuilder(), reader.WorksheetName);
        } while (reader.NextResult());
    }
}