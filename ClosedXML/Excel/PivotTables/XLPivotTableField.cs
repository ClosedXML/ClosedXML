using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

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
    private readonly XLPivotTable _pivotTable;
    private readonly List<XLPivotFieldItem> _items = new();

    public XLPivotTableField(XLPivotTable pivotTable)
    {
        _pivotTable = pivotTable;
        ShowAll = false; // The XML default value is true, but Excel always has false, so let's follow Excel.
        Subtotals = new HashSet<XLSubtotalFunction> { XLSubtotalFunction.Automatic };
    }

    internal XLPivotTable PivotTable => _pivotTable;

    /// <summary>
    /// Pivot field item, doesn't contain value, only indexes to <see cref="XLPivotCache"/> shared items.
    /// </summary>
    internal IReadOnlyList<XLPivotFieldItem> Items => _items;

    #region XML attributes

    /// <summary>
    /// Custom name of the field.
    /// </summary>
    /// <remarks>
    /// [MS-OI29500] Office requires @name to be unique for non-OLAP PivotTables. Ignored by data
    /// fields that use <see cref="XLPivotDataField.DataFieldName"/>.
    /// </remarks>
    internal string? Name { get; set; }

    /// <summary>
    /// If the value is set, the field must also be in <c>rowFields</c>/<c>colFields</c>/
    /// <c>pageFields</c>/<c>dataFields</c> collection in the pivot table part (otherwise Excel
    /// will consider it a corrupt file).
    /// </summary>
    /// <remarks>
    /// [MS-OI29500] In Office, axisValues shall not be used for the axis attribute.
    /// </remarks>
    internal XLPivotAxis? Axis { get; set; }

    /// <summary>
    /// Is this field a data field (i.e. it is referenced the <c>pivotTableDefinition.
    /// dataFields</c>)? Excel will crash, unless these two things both set correctly.
    /// </summary>
    internal bool DataField { get; set; } = false;

    internal string SubtotalCaption { get; set; } = string.Empty;

    internal bool ShowDropDowns { get; init; } = true;

    internal bool HiddenLevel { get; init; } = false;

    internal string? UniqueMemberProperty { get; init; }

    internal bool Compact { get; set; } = true;

    /// <summary>
    /// Are all items expanded?
    /// </summary>
    internal bool AllDrilled { get; set; } = false;

    internal uint? NumberFormatId { get; init; }

    internal bool Outline { get; set; } = true;

    internal bool SubtotalTop { get; set; } = true;

    internal bool DragToRow { get; init; } = true;

    internal bool DragToColumn { get; init; } = true;

    internal bool MultipleItemSelectionAllowed { get; set; } = false;

    internal bool DragToPage { get; init; } = true;

    internal bool DragToData { get; init; } = true;

    internal bool DragOff { get; init; } = true;

    /// <summary>
    /// A flag that indicates whether to show all items for this field.
    /// </summary>
    internal bool ShowAll { get; set; } = true;

    /// <summary>
    /// Insert empty row below every item if the field is row/column axis. The last field in axis
    /// doesn't add extra item at the end. If multiple fields in axis have extra item, only once
    /// blank row is inserted.
    /// </summary>
    internal bool InsertBlankRow { get; set; } = false;

    internal bool ServerField { get; init; } = false;

    internal bool InsertPageBreak { get; set; } = false;

    internal bool AutoShow { get; init; } = false;

    internal bool TopAutoShow { get; init; } = true;

    internal bool HideNewItems { get; init; } = false;

    internal bool MeasureFilter { get; init; } = false;

    internal bool IncludeNewItemsInFilter { get; set; } = false;

    internal uint ItemPageCount { get; init; } = 10;

    internal XLPivotSortType SortType { get; set; } = XLPivotSortType.Default;

    internal bool? DataSourceSort { get; init; }

    internal bool NonAutoSortDefault { get; init; } = false;

    internal uint? RankBy { get; init; }

    /// <summary>
    /// Subtotal functions represented in XML. It's kind of convoluted mess, because
    /// it represents three possible results:
    /// <list type="bullet">
    ///   <item>None - Collection is empty.</item>
    ///   <item>Automatic - Collection contains only <see cref="XLSubtotalFunction.Automatic"/>.</item>
    ///   <item>Custom - Collection contains subtotal functions other than <see cref="XLSubtotalFunction.Automatic"/>.
    ///       The <see cref="XLSubtotalFunction.Automatic"/> is ignored in that case, even if it is present.</item>
    /// </list>.
    /// </summary>
    /// <remarks>
    /// Excel requires that pivot field contains a item if and only if  there is a declared subtotal function.
    /// The subtotal items must be kept at the end of the <see cref="_items"/>, otherwise Excel will try to repair
    /// workbook.
    /// </remarks>
    internal HashSet<XLSubtotalFunction> Subtotals { get; init; }

    internal bool ShowPropCell { get; init; } = false;

    internal bool ShowPropTip { get; init; } = false;

    internal bool ShowPropAsCaption { get; init; } = false;

    internal bool DefaultAttributeDrillState { get; init; } = false;

    /// <summary>
    /// Are item labels on row/column axis repeated for each nested item?
    /// </summary>
    /// <remarks>
    /// Also called <c>FillDownLabels</c>. Attribute is ignored if both the <see cref="Compact"/>
    /// and the <see cref="Outline"/> are <c>true</c>. Attribute is ignored if the field is not on
    /// the <see cref="XLPivotTable.RowAxis"/> or the <see cref="XLPivotTable.ColumnAxis"/>.
    /// </remarks>
    internal bool RepeatItemLabels { get; set; } = false;

    #endregion XML attributes

    internal bool Collapsed
    {
        get => !AllDrilled;
        set => AllDrilled = !value;
    }

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

    internal void AddSubtotal(XLSubtotalFunction value)
    {
        Subtotals.Add(value);
        var subtotalItemType = GetItemTypeForSubtotal(value);;
        _items.Add(new XLPivotFieldItem(this, null) { ItemType = subtotalItemType });
    }

    internal void RemoveSubtotal(XLSubtotalFunction value)
    {
        Subtotals.Remove(value);
        var subtotalItemType = GetItemTypeForSubtotal(value);
        _items.RemoveAll(item => item.ItemType == subtotalItemType);
    }

    internal void SetLayout(XLPivotLayout value)
    {
        switch (value)
        {
            case XLPivotLayout.Compact:
                Outline = false;
                Compact = true;
                break;
            case XLPivotLayout.Outline:
                Compact = false;
                Outline = true;
                break;
            case XLPivotLayout.Tabular:
                Compact = false;
                Outline = false;
                break;
            default:
                throw new UnreachableException();
        }
    }

    internal XLPivotFieldItem GetOrAddItem(XLCellValue value)
    {
        var index = _pivotTable.GetFieldIndex(this);
        var cache = _pivotTable.PivotCache;
        var cacheValues = cache.GetFieldValues(index);
        var sharedItemIndex = cacheValues.GetOrAddSharedItem(value);

        // Excel tries to repair workbook, when there are duplicates in pivotFields.Items
        // therefore add only if necessary
        var existingItem = _items.FirstOrDefault(x => x.ItemIndex == sharedItemIndex);
        if (existingItem is not null)
            return existingItem;

        var newItem = new XLPivotFieldItem(this, sharedItemIndex);
        _items.Add(newItem);
        return newItem;
    }

    private static XLPivotItemType GetItemTypeForSubtotal(XLSubtotalFunction value)
    {
        var subtotalItemType = value switch
        {
            XLSubtotalFunction.Automatic => XLPivotItemType.Default,
            XLSubtotalFunction.Sum => XLPivotItemType.Sum,
            XLSubtotalFunction.Count => XLPivotItemType.CountA,
            XLSubtotalFunction.Average => XLPivotItemType.Avg,
            XLSubtotalFunction.Minimum => XLPivotItemType.Min,
            XLSubtotalFunction.Maximum => XLPivotItemType.Max,
            XLSubtotalFunction.Product => XLPivotItemType.Product,
            XLSubtotalFunction.CountNumbers => XLPivotItemType.Count,
            XLSubtotalFunction.StandardDeviation => XLPivotItemType.StdDev,
            XLSubtotalFunction.PopulationStandardDeviation => XLPivotItemType.StdDevP,
            XLSubtotalFunction.Variance => XLPivotItemType.Var,
            XLSubtotalFunction.PopulationVariance => XLPivotItemType.VarP,
            _ => throw new UnreachableException()
        };
        return subtotalItemType;
    }
}
