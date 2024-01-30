using System.Collections.Generic;

namespace ClosedXML.Excel;

/// <summary>
/// One field in a <see cref="XLPivotTable"/>. Pivot table must contain field for each entry of
/// pivot cache and both are accessed through same index. Pivot field contains items, which are
/// cache field values referenced anywhere in the pivot table (e.g. caption, axis value ect.).
/// </summary>
/// <remarks>
/// See <em>[OI-29500] 18.10.1.69 pivotField(PivotTable Field)</em> for details.
/// </remarks>
internal class XLPivotTableField
{
    private readonly List<XLPivotFieldItem> _items = new();

    /// <summary>
    /// Pivot field item, doesn't contain value, only indexes to <see cref="XLPivotCache"/> shared items.
    /// </summary>
    internal IReadOnlyList<XLPivotFieldItem> Items => _items;

    /// <summary>
    /// Custom name of the field.
    /// </summary>
    /// <remarks>
    /// [MS-OI29500] Office requires @name to be unique for non-OLAP PivotTables.
    /// </remarks>
    internal string? Name { get; init; }

    /// <summary>
    /// </summary>
    /// <remarks>
    /// [MS-OI29500] In Office, axisValues shall not be used for the axis attribute.
    /// </remarks>
    internal XLPivotAxis? Axis { get; init; }

    internal bool DataField { get; init; } = false;

    internal string? SubtotalCaption { get; init; }

    internal bool ShowDropDowns { get; init; } = true;

    internal bool HiddenLevel { get; init; } = false;

    internal string? UniqueMemberProperty { get; init; }

    internal bool Compact { get; init; } = true;

    internal bool AllDrilled { get; init; } = false;

    internal uint? NumberFormatId { get; init; }

    internal bool Outline { get; init; } = true;

    internal bool SubtotalTop { get; init; } = true;

    internal bool DragToRow { get; init; } = true;

    internal bool DragToColumn { get; init; } = true;

    internal bool MultipleItemSelectionAllowed { get; init; } = false;

    internal bool DragToPage { get; init; } = true;

    internal bool DragToData { get; init; } = true;

    internal bool DragOff { get; init; } = true;

    internal bool ShowAll { get; init; } = true;

    internal bool InsertBlankRow { get; init; } = false;

    internal bool ServerField { get; init; } = false;

    internal bool InsertPageBreak { get; init; } = false;

    internal bool AutoShow { get; init; } = false;

    internal bool TopAutoShow { get; init; } = true;

    internal bool HideNewItems { get; init; } = false;

    internal bool MeasureFilter { get; init; } = false;

    internal bool IncludeNewItemsInFilter { get; init; } = false;

    internal uint ItemPageCount { get; init; } = 10;

    internal XLPivotSortType SortType { get; init; } = XLPivotSortType.Default;

    internal bool? DataSourceSort { get; init; }

    internal bool NonAutoSortDefault { get; init; } = false;

    internal uint? RankBy { get; init; }

    internal bool DefaultSubtotal { get; init; } = true;

    internal bool SumSubtotal { get; init; } = false;

    internal bool CountASubtotal { get; init; } = false;

    internal bool AvgSubtotal { get; init; } = false;

    internal bool MaxSubtotal { get; init; } = false;

    internal bool MinSubtotal { get; init; } = false;

    internal bool ProductSubtotal { get; init; } = false;

    internal bool CountSubtotal { get; init; } = false;

    internal bool StdDevSubtotal { get; init; } = false;

    internal bool StdDevPSubtotal { get; init; } = false;

    internal bool VarSubtotal { get; init; } = false;

    internal bool VarPSubtotal { get; init; } = false;

    internal bool ShowPropCell { get; init; } = false;

    internal bool ShowPropTip { get; init; } = false;

    internal bool ShowPropAsCaption { get; init; } = false;

    internal bool DefaultAttributeDrillState { get; init; } = false;

    /// <summary>
    /// Add an item when it is used anywhere in the pivot table.
    /// </summary>
    /// <param name="item">Item to add.</param>
    /// <returns>Index of added item.</returns>
    internal uint AddItem(XLPivotFieldItem item)
    {
        var index = _items.Count;
        _items.Add(item);
        return (uint)index;
    }
}
