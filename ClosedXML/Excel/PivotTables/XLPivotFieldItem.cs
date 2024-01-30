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
    private readonly XLPivotTableField _pivotField;

    internal XLPivotFieldItem(XLPivotTableField pivotField, uint? itemIndex)
    {
        // TODO: Check that index is in shared items of cached fields.
        _pivotField = pivotField;
        ItemIndex = itemIndex;
    }

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
    internal bool Hidden { get; init; } = false;

    /// <summary>Flag indicating that the item has a character value/</summary>
    /// <remarks>Allowed for OLAP pivot tables only.</remarks>
    internal bool ValueIsString { get; init; } = false;

    /// <remarks>Allowed for non-OLAP pivot tables only.</remarks>
    internal bool HideDetails { get; init; } = true;

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
    /// </summary>
    internal uint? ItemIndex { get; }

    /// <remarks>Allowed for OLAP pivot tables only.</remarks>
    internal bool DrillAcrossAttributes { get; init; }

    /// <remarks>Allowed for OLAP pivot tables only.</remarks>
    internal bool IsExpanded { get; init; }
}
