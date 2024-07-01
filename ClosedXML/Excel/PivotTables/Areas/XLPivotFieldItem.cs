using System;
using System.Diagnostics.CodeAnalysis;

namespace ClosedXML.Excel;

/// <summary>
/// Representation of item (basically one value of a field). Each value used somewhere in pivot
/// table (e.g. data area, row/column labels and so on) must have an entry here. By itself, it
/// doesn't contain values, it only references shared items of the field in the
/// <see cref="XLPivotCache"/>.
/// </summary>
/// <remarks>
/// [OI29500] 18.10.1.45 item (PivotTable Field Item). 
/// </remarks>
internal class XLPivotFieldItem
{
    private readonly XLPivotTable _pivotTable;
    private readonly XLPivotTableField _pivotField;

    internal XLPivotFieldItem(XLPivotTableField pivotField, int? itemIndex)
    {
        if (itemIndex.HasValue && itemIndex.Value < 0)
            throw new ArgumentOutOfRangeException(nameof(itemIndex));

        _pivotField = pivotField;
        _pivotTable = pivotField.PivotTable;
        ItemIndex = itemIndex;
    }

    #region XML attributes

    /// <summary>
    /// If present, must be unique within the containing field items.
    /// </summary>
    internal string? ItemUserCaption { get; init; }

    internal XLPivotItemType ItemType { get; init; } = XLPivotItemType.Data;

    /// <summary>
    /// <para>
    /// Flag indicating that the the item hidden. Used for <see cref="XLPivotPageField"/>. When
    /// item field is a page field, the hidden flag mean unselected values in the page filter.
    /// Non-hidden values are selected in the filter.
    /// </para>
    /// <para>
    /// Allowed for non-OLAP pivot tables only.
    /// </para>
    /// </summary>
    internal bool Hidden { get; set; } = false;

    /// <summary>
    /// Flag indicating that the item has a character value.
    /// </summary>
    /// <remarks>Allowed for OLAP pivot tables only.</remarks>
    internal bool ValueIsString { get; init; } = false;

    /// <summary>
    /// Excel uses the <c>sd</c> attribute to indicate whether the item is expanded.
    /// </summary>
    /// <remarks>Allowed for non-OLAP pivot tables only. Spec for the <c>sd</c> had to be patched..</remarks>
    internal bool ShowDetails { get; set; } = true;

    /// <remarks>Allowed for non-OLAP pivot tables only.</remarks>
    internal bool CalculatedMember { get; init; } = false;

    /// <summary>
    /// Item itself is missing from the source data
    /// </summary>
    /// <remarks>Allowed for non-OLAP pivot tables only.</remarks>
    internal bool Missing { get; init; } = false;

    /// <remarks>Allowed for OLAP pivot tables only.</remarks>
    internal bool ApproximatelyHasChildren { get; init; } = false;

    /// <summary>
    /// Index to an item in the sharedItems of the field. The index must be unique in containing field items. When <see cref="ItemType"/> is <see cref="XLPivotItemType.Data"/>, it must be set.
    /// Never negative.
    /// </summary>
    internal int? ItemIndex { get; }

    /// <remarks>Allowed for OLAP pivot tables only.</remarks>
    internal bool DrillAcrossAttributes { get; init; } = true;

    /// <summary>
    /// Attributes <c>sd</c> (show detail) and <c>d</c> (detail) were swapped in spec, fixed by OI29500.
    /// A flag that indicates whether details are hidden for this item?
    /// </summary>
    /// <remarks><c>d</c> attribute. Allowed for OLAP pivot tables only.</remarks>
    internal bool Details { get; init; }

    #endregion XML attributes

    [MemberNotNullWhen(true, nameof(ItemIndex))]
    private bool ValueIsData => ItemType == XLPivotItemType.Data;

    /// <summary>
    /// Get value of an item from cache or null if not data item.
    /// </summary>
    internal XLCellValue? GetValue()
    {
        if (!ValueIsData)
            return null;

        var fieldIndex = _pivotTable.GetFieldIndex(_pivotField);
        var sharedItems = _pivotTable.PivotCache.GetFieldSharedItems(fieldIndex);
        var itemIndex = ItemIndex.Value;
        return sharedItems[checked((uint)itemIndex)];
    }
}
