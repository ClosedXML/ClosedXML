namespace ClosedXML.Excel;

/// <summary>
/// An API to modify a differential format of a workbook object (table, conditional format ect.).
/// </summary>
internal interface IXLDifferentialFormat
{
    /// <summary>
    /// A property to modify font properties of the differential format.
    /// </summary>
    IXLFontFormat<IXLDifferentialFormat> Font { get; }
}
