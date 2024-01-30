namespace ClosedXML.Excel;

/// <summary>
/// A field displayed in the filters part of a pivot table.
/// </summary>
internal class XLPivotPageField
{
    internal XLPivotPageField(int field)
    {
        Field = field;
    }

    /// <summary>
    /// Field index to <see cref="XLPivotTable.PivotFields"/>. Can't contain 'data' field <c>-2</c>.
    /// </summary>
    internal int Field { get; }

    /// <summary>
    /// If a single item is selected, item index. Null, if nothing selected or multiple selected.
    /// Multiple selected values are indicated directly in <see cref="XLPivotTableField.Items"/>
    /// through <see cref="XLPivotFieldItem.Hidden"/> flags. Items that are not selected are hidden,
    /// rest isn't.
    /// </summary>
    internal uint? ItemIndex { get; init; }

    // OLAP
    internal int? HierarchyIndex { get; init; }

    // OLAP
    internal string? HierarchyUniqueName { get; init; }

    // OLAP
    internal string? HierarchyDisplayName { get; init; }
}
