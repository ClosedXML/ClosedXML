// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel;

/// <summary>
/// A fluent API for styling a <see cref="IXLPivotValue"/> field of a <see cref="IXLPivotTable"/>.
/// </summary>
public interface IXLPivotValueStyleFormat : IXLPivotStyleFormat
{
    /// <summary>
    /// Adds a further limitation so the <see cref="IXLPivotStyleFormat.Style"/> is only applied to cells in a pivot table
    /// that also belong to the <paramref name="field"/> (label or data).
    /// </summary>
    /// <param name="field">Only cells in a pivot table under this field will be styled.</param>
    IXLPivotValueStyleFormat AndWith(IXLPivotField field);

    /// <summary>
    /// Adds a further limitation so the <see cref="IXLPivotStyleFormat.Style"/> is only applied to cells in a pivot table
    /// that also belong to the <paramref name="field"/> data cells. The cell values also must satisfy the <paramref name="predicate"/>.
    /// </summary>
    /// <remarks>
    /// The pivot style is bound by the field index in a pivot table, not field value. E.g. if field values
    /// are Jan, Feb and the predicate marks Feb (offset 1) = second field (Feb) will be highlighted.
    /// If user later reverses order in Excel to Feb, Jan, the style would still apply to the second value - Jan.
    /// </remarks>
    /// <param name="field">Only cells in a pivot table under this field will be styled.</param>
    /// <param name="predicate">A predicate to determine which index of the field should be styled.</param>
    IXLPivotValueStyleFormat AndWith(IXLPivotField field, Predicate<XLCellValue> predicate);

    IXLPivotValueStyleFormat ForValueField(IXLPivotValue valueField);
}
