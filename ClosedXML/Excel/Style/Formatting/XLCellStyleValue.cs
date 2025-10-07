namespace ClosedXML.Excel.Formatting;

/// <summary>
/// <para>
/// A cell style available in the workbook.
/// </para>
/// <para>
/// The style is intended to be mutable, unlike cell format. Changing a style aspect (name or size)
/// shouldn't require update of all cells that use it. The <see cref="XLCellFormatValue"/> links to
/// the style through <see cref="XLCellFormatValue.CellStyleId"/> and the <see cref="XLWorkbookStyles.CellStyles"/>
/// collection.
/// </para>
/// </summary>
internal class XLCellStyleValue
{
    /// <summary>
    /// A unique name of the style.
    /// </summary>
    public required string Name { get; init; }

    public required BuiltInStyleValues? BuiltInStyle { get; init; }

    /// <summary>
    /// Is style hidden in the UI?
    /// </summary>
    public required bool Hidden { get; init; }

    public required string NumberFormat { get; init; }

    public required XLAlignmentFormatValue Alignment { get; init; }

    public required XLProtectionFormatValue Protection { get; init; }

    public required XLFontFormatValue Font { get; init; }

    public required XLFillFormatValue Fill { get; init; }

    public required XLBorderFormatValue Border { get; init; }

    /// <summary>
    /// Format components that are decided by the style. Specified components should have
    /// a non-null value.
    /// </summary>
    public required CellFormatComponents IncludedComponents { get; init; }
}
