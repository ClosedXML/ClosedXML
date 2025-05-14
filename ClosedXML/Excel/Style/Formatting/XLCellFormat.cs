namespace ClosedXML.Excel.Formatting;

/// <summary>
/// <para>
/// A master formatting record that determines a direct formatting of a cell/column/row. The final
/// formatting used to render a cell is determined by composition of multiple master formatting
/// records at multiple levels. The least specific is the default one, unrelated to a workbook.
/// Next level is a workbook formatting record, represented by normal style. Next is column or row
/// one and the most specific one is in the cell.
/// </para>
/// <para>
/// Master formatting record has optional properties. The unspecified properties are set through
/// formatting records composition. The <see cref="XLStyleKey"/> (or its reference form
/// <see cref="XLStyleValue"/>) has everything specified, because it is the final formatting of a
/// cell.
/// </para>
/// </summary>
internal record XLCellFormat
{
    public required string? NumberFormat { get; init; }

    public required XLAlignmentFormat? Alignment { get; init; }

    public required XLProtectionFormat? Protection { get; init; }

    public required XLFontFormat? Font { get; init; }

    public required XLFillFormat? Fill { get; init; }

    public required XLBorderFormat? Border { get; init; }

    /// <summary>
    /// A cell style that was originally used to create this format.
    /// </summary>
    public required XLCellStyle? CellStyle {get; init; }

    public required bool QuotePrefix { get; init; }

    public required bool PivotButton { get; init; }

    /// <summary>
    /// <para>
    /// Format components that should be updated when the original <see cref="CellStyle"/> is
    /// changed. The format is immutable, so the change is actually creation of derived format and
    /// usage replacement.
    /// </para>
    /// <para>
    /// Only used if format was created from a style.
    /// </para>
    /// <para>
    /// <example>
    /// User stylizes a cell with an <em>Input</em> style that specifies a font, fill and border.
    /// User then changes size of a text in a cell, thus the cell format now contains a different
    /// font format. If the style <em>Input</em> changes background, the cell format should now use
    /// a new background. But if style <em>Input</em> changes font, it shouldn't be reflected in
    /// the format, because it was explicitly set to a different value from a style.
    /// </example>
    /// </para>
    /// </summary>
    public required CellFormatComponents StyleComponents { get; init; }
}
