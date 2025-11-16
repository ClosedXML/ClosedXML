namespace ClosedXML.Excel.Formatting;

/// <summary>
/// A differential format.
/// </summary>
internal record XLDxfValue
{
    internal static readonly XLDxfValue Empty = new()
    {
        NumberFormat = null,
        Font = XLDifferentialFontValue.Empty,
        Fill = null,
        Alignment = null,
        Border = null,
        Protection = null
    };

    public required string? NumberFormat { get; init; }

    public required XLDifferentialFontValue Font { get; init; }

    public required XLFillFormatValue? Fill { get; init; }

    public required XLAlignmentFormatValue? Alignment { get; init; }

    public required XLBorderFormatValue? Border { get; init; }

    public required XLProtectionFormatValue? Protection { get; init; }
}
