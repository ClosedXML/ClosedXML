namespace ClosedXML.Excel.Formatting;

/// <summary>
/// Border line properties.
/// </summary>
internal readonly record struct XLBorderLine(XLColor Color, XLBorderStyleValues Style)
{
    internal static readonly XLBorderLine None = new(XLColor.Auto, XLBorderStyleValues.None);
}
