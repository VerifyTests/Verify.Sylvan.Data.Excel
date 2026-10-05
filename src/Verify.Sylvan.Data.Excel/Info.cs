class Info
{
    public required IEnumerable<string> SheetNames { get; init; }

    // A hidden sheet has a csv as any other, so this is what says which are hidden. Null, so left
    // out, for a workbook with none, and for a reader that was passed in, which does not say.
    public IReadOnlyList<string>? HiddenSheets { get; init; }
}
