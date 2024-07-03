// Keep this file CodeMaid organised and cleaned
namespace ClosedXML.Excel;

/// <summary>
/// A interface for styling various parts of a pivot table, e.g. the whole table, specific
/// area or just a field. Use <see cref="IXLPivotTable.StyleFormats"/> and <see cref="IXLPivotField.StyleFormats"/>
/// to access it.
/// </summary>
public interface IXLPivotStyleFormat
{
    /// <summary>
    /// To what part of the pivot table part will the style apply to.
    /// </summary>
    XLPivotStyleFormatElement AppliesTo { get; }

    /// <summary>
    /// The differential style of the part.
    /// </summary>
    /// <remarks>
    /// The final displayed style is done by composing all differential styles that overlap the element.
    /// </remarks>
    IXLStyle Style { get; set; }
}
