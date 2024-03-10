#nullable disable
using System;
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
    private readonly FieldIndex _index;

    internal XLPivotTablePageField(XLPivotTable pivotTable, FieldIndex index)
    {
        _pivotTable = pivotTable;
        _index = index;
    }

    public string SourceName
    {
        get
        {
            if (_index.IsDataField)
                throw new InvalidOperationException("Filter field can't contain data field.");

            return _pivotTable.PivotCache.FieldNames[_index];
        }
    }

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
        var fieldItem = GetField().GetOrAddItem(value);
        fieldItem.Hidden = false;
        return this;
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
    public int Offset => _index;

    private XLPivotTableField GetField()
    {
        if (_index.IsDataField)
            throw new InvalidOperationException("Can't set this property on a data field.");

        return _pivotTable.PivotFields[_index];
    }
}
