namespace ClosedXML.Excel.Formatting;

/// <summary>
/// A differential format.
/// </summary>
internal record XLDxfValue
{
    public required string? NumberFormat { get; init; }

    public required XLFontFormatValue? Font { get; init; }

    public required XLFillFormatValue? Fill { get; init; }

    public required XLAlignmentFormatValue? Alignment { get; init; }

    public required XLBorderFormatValue? Border { get; init; }

    public required XLProtectionFormatValue? Protection { get; init; }
}
