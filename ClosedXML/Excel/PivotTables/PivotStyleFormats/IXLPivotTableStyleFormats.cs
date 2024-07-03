// Keep this file CodeMaid organised and cleaned
namespace ClosedXML.Excel;

/// <summary>
/// An API for modifying the pivot table styles that affect whole <see cref="IXLPivotTable"/>.
/// </summary>
public interface IXLPivotTableStyleFormats
{
    /// <summary>
    /// Get style formats of a grand total column in a pivot table (i.e. the right column a pivot table).
    /// </summary>
    IXLPivotStyleFormats ColumnGrandTotalFormats { get; }

    /// <summary>
    /// Get style formats of a grand total row in a pivot table (i.e. the bottom row of a pivot table).
    /// </summary>
    IXLPivotStyleFormats RowGrandTotalFormats { get; }
}
