namespace ClosedXML.Excel;

/// <summary>
/// A description of formatting that should be applied to a <see cref="XLPivotTable"/>.
/// </summary>
internal class XLPivotFormat
{
    internal XLPivotFormat(XLPivotArea pivotArea)
    {
        PivotArea = pivotArea;
    }

    /// <summary>
    /// Pivot area that should be formatted.
    /// </summary>
    internal XLPivotArea PivotArea { get; }

    /// <summary>
    /// Should the formatting (determined by <see cref="DxfStyleValue"/>) be applied or not?
    /// </summary>
    internal XLPivotFormatAction Action { get; init; } = XLPivotFormatAction.Formatting;

    /// <summary>
    /// Differential formatting to apply to the <see cref="PivotArea"/>. It can be empty, e.g. if
    /// <see cref="Action"/> is blank. Empty dxf is represented by <see cref="XLStyle.Default"/>,
    /// until we get better dxf representation.
    /// </summary>
    internal XLStyleValue DxfStyleValue { get; set; } = XLStyle.Default.Value;
}
