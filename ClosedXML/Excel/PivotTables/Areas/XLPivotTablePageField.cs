#nullable disable
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel;

/// <summary>
/// Fluent API for filter fields of a <see cref="XLPivotTable"/>. This class shouldn't contain any
/// state, only logic to change state per API.
/// </summary>
internal class XLPivotTablePageField : IXLPivotField
{
    private readonly XLPivotTable _pivotTable;
    private readonly XLPivotPageField _filterField;

    internal XLPivotTablePageField(XLPivotTable pivotTable, XLPivotPageField filterField)
    {
        _pivotTable = pivotTable;
        _filterField = filterField;
    }

    public string SourceName => _pivotTable.PivotCache.FieldNames[_filterField.Field];

    public string CustomName
    {
        get => GetField().Name;
        set => GetField().Name = value;
    }

    public string SubtotalCaption
    {
        get => GetField().SubtotalCaption;
        set => GetField().SubtotalCaption = value;
    }

    public IReadOnlyCollection<XLSubtotalFunction> Subtotals => GetField().Subtotals;

    public bool IncludeNewItemsInFilter
    {
        get => GetField().IncludeNewItemsInFilter;
        set => GetField().IncludeNewItemsInFilter = value;
    }

    public bool Outline
    {
        get => GetField().Outline;
        set => GetField().Outline = value;
    }

    public bool Compact
    {
        get => GetField().Compact;
        set => GetField().Compact = value;
    }

    public bool? SubtotalsAtTop
    {
        get => GetField().SubtotalTop;
        set => GetField().SubtotalTop = value ?? true;
    }

    public bool RepeatItemLabels
    {
        get => GetField().RepeatItemLabels;
        set => GetField().RepeatItemLabels = value;
    }

    public bool InsertBlankLines
    {
        get => GetField().InsertBlankRow;
        set => GetField().InsertBlankRow = value;
    }

    public bool ShowBlankItems
    {
        get => GetField().ShowAll;
        set => GetField().ShowAll = value;
    }

    public bool InsertPageBreaks
    {
        get => GetField().InsertPageBreak;
        set => GetField().InsertPageBreak = value;
    }

    public bool Collapsed
    {
        get => GetField().Collapsed;
        set => GetField().Collapsed = value;
    }

    public XLPivotSortType SortType
    {
        get => GetField().SortType;
        set => GetField().SortType = value;
    }

    public IXLPivotField SetCustomName(string value)
    {
        CustomName = value;
        return this;
    }

    public IXLPivotField SetSubtotalCaption(string value)
    {
        SubtotalCaption = value;
        return this;
    }

    public IXLPivotField AddSubtotal(XLSubtotalFunction value)
    {
        GetField().AddSubtotal(value);
        return this;
    }

    public IXLPivotField SetIncludeNewItemsInFilter(bool value)
    {
        IncludeNewItemsInFilter = value;
        return this;
    }

    public IXLPivotField SetLayout(XLPivotLayout value)
    {
        GetField().SetLayout(value);
        return this;
    }

    public IXLPivotField SetSubtotalsAtTop(bool value)
    {
        SubtotalsAtTop = value;
        return this;
    }

    public IXLPivotField SetRepeatItemLabels(bool value)
    {
        RepeatItemLabels = value;
        return this;
    }

    public IXLPivotField SetInsertBlankLines(bool value)
    {
        InsertBlankLines = value;
        return this;
    }

    public IXLPivotField SetShowBlankItems(bool value)
    {
        ShowBlankItems = value;
        return this;
    }

    public IXLPivotField SetInsertPageBreaks(bool value)
    {
        InsertPageBreaks = value;
        return this;
    }

    public IXLPivotField SetCollapsed(bool value)
    {
        Collapsed = value;
        return this;
    }

    public IXLPivotField SetSort(XLPivotSortType value)
    {
        SortType = value;
        return this;
    }

    public IReadOnlyList<XLCellValue> SelectedValues
    {
        get
        {
            var shownItems = GetField().Items.Where(i => !i.Hidden);
            var selectedValues = new List<XLCellValue>();
            foreach (var selectedItem in shownItems)
            {
                var selectedValue = selectedItem.GetValue();
                if (selectedValue is not null)
                    selectedValues.Add(selectedValue.Value);
            }

            return selectedValues;
        }
    }

    public IXLPivotField AddSelectedValue(XLCellValue value)
    {
        // Try to keep the original behavior of ClosedXML - it always allows multiple selected items for added values.
        // But it's complete kludge with no sane semantic that will be nuked ASAP.
        var pivotField = GetField();

        var nothingSelected = _filterField.ItemIndex is null && !pivotField.MultipleItemSelectionAllowed;
        if (nothingSelected)
        {
            var fieldItem = pivotField.GetOrAddItem(value);
            _filterField.ItemIndex = (uint?)fieldItem.ItemIndex;
            return this;
        }

        var oneItemSelected = _filterField.ItemIndex is not null && !pivotField.MultipleItemSelectionAllowed;
        if (oneItemSelected)
        {
            // Switch to multiple
            pivotField.MultipleItemSelectionAllowed = true;
            foreach (var item in pivotField.Items.Where(x => x.ItemType == XLPivotItemType.Data))
                item.Hidden = true;

            var selectedItem = pivotField.Items.Single(i => i.ItemIndex == _filterField.ItemIndex);
            selectedItem.Hidden = false;
            _filterField.ItemIndex = null;
            var fieldItem = pivotField.GetOrAddItem(value);
            fieldItem.Hidden = false;
            return this;
        }
        else
        {
            // Add another item to selected item filters.
            var fieldItem = pivotField.GetOrAddItem(value);
            fieldItem.Hidden = false;
            return this;
        }
    }

    public IXLPivotField AddSelectedValues(IEnumerable<XLCellValue> values)
    {
        foreach (var value in values)
            AddSelectedValue(value);

        return this;
    }

    public IXLPivotFieldStyleFormats StyleFormats => new XLPivotFieldStyleFormats(this);
    public bool IsOnRowAxis => false;
    public bool IsOnColumnAxis => false;
    public bool IsInFilterList => true;
    public int Offset => _filterField.Field;

    private XLPivotTableField GetField()
    {
        return _pivotTable.PivotFields[_filterField.Field];
    }
}
