namespace ClosedXML.Excel.Formatting;

/// <summary>
/// A differential format.
/// </summary>
internal record XLDifferentialFormat
{
    public required string? NumberFormat { get; init; }

    public required XLFontFormat? Font { get; init; }

    public required XLFillFormat? Fill { get; init; }

    public required XLAlignmentFormat? Alignment { get; init; }

    public required XLBorderFormat? Border { get; init; }

    public required XLProtectionFormat? Protection { get; init; }
}
