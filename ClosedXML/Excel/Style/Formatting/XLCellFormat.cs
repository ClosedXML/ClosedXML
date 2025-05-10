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
internal readonly record struct XLCellFormat
{
    public XLFontFormat? Font { get; init; }

    public XLBorderFormat? Border { get; init; }

    public XLFillFormat? Fill { get; init; }

    // TODO: Add remaining properties. For now only font
}
