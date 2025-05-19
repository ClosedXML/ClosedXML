namespace ClosedXML.Excel.Formatting;

/// <summary>
/// <para>
/// A cell style available in the workbook.
/// </para>
/// <para>
/// Cell style isn't actually used for formatting. A <see cref="XLCellFormat"/> is created from
/// a style and the cell format is then used to format.
/// </para>
/// </summary>
internal class XLCellStyle
{
    /// <summary>
    /// Name of the style.
    /// </summary>
    public required string Name { get; init; }

    public required BuiltInStyleValues? BuiltInStyle { get; init; }

    /// <summary>
    /// Is style hidden in the UI?
    /// </summary>
    public required bool Hidden { get; init; }

    public required string? NumberFormat { get; init; }

    public required XLAlignmentFormat? Alignment { get; init; }

    public required XLProtectionFormat? Protection { get; init; }

    public required XLFontFormat? Font { get; init; }

    public required XLFillFormat? Fill { get; init; }

    public required XLBorderFormat? Border { get; init; }

    public required CellFormatComponents ApplyComponents { get; init; }
}
