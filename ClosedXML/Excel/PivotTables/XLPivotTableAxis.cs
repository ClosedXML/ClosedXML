using System.Collections.Generic;

namespace ClosedXML.Excel;

/// <summary>
/// A description of one axis (<see cref="XLPivotTable.RowAxis"/>/<see cref="XLPivotTable.ColumnAxis"/>)
/// of a <see cref="XLPivotTable"/>. It consists of fields in a specific order and values that make up
/// individual rows/columns of the axis.
/// </summary>
/// <remarks>
/// [ISO-29500] 18.10.1.17 colItems (Column Items), 18.10.1.84 rowItems (Row Items).
/// </remarks>
internal class XLPivotTableAxis
{
    private readonly List<FieldIndex> _fields = new();

    /// <summary>
    /// Value of one row/column in an axis.
    /// </summary>
    private readonly List<XLPivotFieldAxisItem> _axisItems = new();

    internal XLPivotTableAxis(XLPivotTable pivotTable)
    {
        PivotTable = pivotTable;
    }

    /// <summary>
    /// Pivot table this axis belongs to.
    /// </summary>
    internal XLPivotTable PivotTable { get; }

    /// <summary>
    /// A list of fields to displayed on the axis. It determines which fields and in what order
    /// should the fields be displayed.
    /// </summary>
    internal IReadOnlyList<FieldIndex> Fields => _fields;

    /// <summary>
    /// Individual row/column parts of the axis.
    /// </summary>
    internal IReadOnlyList<XLPivotFieldAxisItem> Items => _axisItems;

    /// <summary>
    /// Add field to the axis.
    /// </summary>
    internal void AddField(FieldIndex fieldIndex)
    {
        _fields.Add(fieldIndex);
    }

    /// <summary>
    /// Add a row/column axis values (i.e. values visible on the axis).
    /// </summary>
    internal void AddItem(XLPivotFieldAxisItem axisItem)
    {
        _axisItems.Add(axisItem);
    }
}
