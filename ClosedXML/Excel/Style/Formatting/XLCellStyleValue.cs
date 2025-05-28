namespace ClosedXML.Excel.Formatting;

/// <summary>
/// <para>
/// A cell style available in the workbook.
/// </para>
/// <para>
/// Cell style isn't actually used for formatting. A <see cref="XLCellFormatValue"/> is created from
/// a style and the cell format is then used to format.
/// </para>
/// </summary>
internal class XLCellStyleValue
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

    public required XLAlignmentFormatValue? Alignment { get; init; }

    public required XLProtectionFormatValue? Protection { get; init; }

    public required XLFontFormatValue? Font { get; init; }

    public required XLFillFormatValue? Fill { get; init; }

    public required XLBorderFormatValue? Border { get; init; }

    public required CellFormatComponents ApplyComponents { get; init; }
}
