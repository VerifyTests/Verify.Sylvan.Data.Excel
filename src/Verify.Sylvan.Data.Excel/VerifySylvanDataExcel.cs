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

        // A reader does not say which sheets are hidden, only whether it reads them. So the names
        // it has when it does are compared with the names it has when it does not.
        var all = SheetNames(bytes, type, true);
        var visible = SheetNames(bytes, type, false);

        // A hidden sheet is read as any other, which a reader does not do unless it is asked to.
        // Options that were passed in are used as they are, so they can still leave them out.
        var options = settings.GetExcelDataReaderOptions() ??
                      new ExcelDataReaderOptions
                      {
                          ReadHiddenWorksheets = true
                      };
        bytes.Position = 0;
        using var reader = ExcelDataReader.Create(bytes, type, options);
        return Convert(reader, settings, BuildSource(bytes, type, settings), all, HiddenSheets(bytes, type, all, visible));
    }

    // In the order of the sheets. A reader takes a sheet that only code can unhide to be one that
    // is not hidden at all, so for an xlsx, which says so in a part that can be read, those are
    // added. An xlsb and an xls say so in binary records, and there such a sheet is not named.
    static List<string>? HiddenSheets(MemoryStream bytes, ExcelWorkbookType type, List<string> all, List<string> visible)
    {
        var names = new HashSet<string>(all.Except(visible));
        if (type == ExcelWorkbookType.ExcelXml)
        {
            names.UnionWith(SheetsWithAState(bytes));
        }

        var hidden = all.Where(names.Contains).ToList();
        if (hidden.Count == 0)
        {
            return null;
        }

        return hidden;
    }

    // The sheets xl/workbook.xml gives a state of hidden or veryHidden
    static IEnumerable<string> SheetsWithAState(MemoryStream bytes)
    {
        using var copy = new MemoryStream(bytes.ToArray());
        using var archive = new ZipArchive(copy, ZipArchiveMode.Read);
        var entry = archive.GetEntry("xl/workbook.xml");
        if (entry is null)
        {
            return [];
        }

        using var stream = entry.Open();
        return XDocument.Load(stream)
            .Descendants()
            .Where(_ => _.Name.LocalName == "sheet" &&
                        _.Attribute("state")?.Value is "hidden" or "veryHidden")
            .Select(_ => _.Attribute("name")!.Value)
            .ToList();
    }

    static List<string> SheetNames(MemoryStream bytes, ExcelWorkbookType type, bool readHidden)
    {
        // A copy, since a reader closes the stream it is given
        using var copy = new MemoryStream(bytes.ToArray());
        var options = new ExcelDataReaderOptions
        {
            ReadHiddenWorksheets = readHidden
        };
        using var reader = ExcelDataReader.Create(copy, type, options);
        return reader.WorksheetNames.ToList();
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

    // A reader passed in directly has no workbook to snapshot: the bytes it reads are not to hand.
    // Nor is there any telling which sheets it was not asked to read, so the sheets it has are the
    // pages.
    static ConversionResult Convert(ExcelDataReader reader, IReadOnlyDictionary<string, object> settings) =>
        Convert(reader, settings, null, reader.WorksheetNames.ToList(), null);

    // A sheet is a page, numbered by where it is among all the sheets of the workbook, hidden or
    // not, so PagesToInclude limits the csv files as it limits the pages of any other document.
    static ConversionResult Convert(ExcelDataReader reader, IReadOnlyDictionary<string, object> settings, Target? source, List<string> pages, List<string>? hidden)
    {
        var info = new Info
        {
            SheetNames = reader.WorksheetNames,
            HiddenSheets = hidden,
        };

        // Reading the sheets is the expensive part, so it is skipped when the csv files are
        // excluded, by ExcludeDerivedTargets or by ExcludeTargets
        List<Target> sheets = [];
        if (!settings.IsDerivedTargetExcluded("csv"))
        {
            var options = settings.GetCsvDataWriterOptions();
            sheets.AddRange(Convert(reader, options, pages, settings));
        }

        // Verify names the csv files relative to the target that was converted, and with a source
        // compares it first and tells the diff tool the csv files came from it
        return new(info, source, derived: sheets);
    }

    static IEnumerable<Target> Convert(ExcelDataReader reader, CsvDataWriterOptions? options, List<string> pages, IReadOnlyDictionary<string, object> settings)
    {
        // No name is no sheet: a workbook with none to read, or a reader moved past its last
        if (reader.WorksheetName is null)
        {
            yield break;
        }

        do
        {
            // A sheet that is not read is passed over, which costs nothing
            var page = pages.IndexOf(reader.WorksheetName!) + 1;
            if (!settings.IsPageIncluded(page))
            {
                continue;
            }

            using var writer = new StringWriter();
            using var csvWriter = CsvDataWriter.Create(writer, options);
            csvWriter.Write(reader);
            // Named after its sheet even when it is the only one, so that a second sheet adds a
            // file rather than renaming the first
            yield return new("csv", writer.GetStringBuilder(), reader.WorksheetName);
        } while (reader.NextResult());
    }
}