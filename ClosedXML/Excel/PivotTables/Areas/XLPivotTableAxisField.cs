#nullable disable
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel;

/// <summary>
/// A fluent API for one field in <see cref="XLPivotTableAxis"/>, either
/// <see cref="XLPivotTable.RowLabels"/> or <see cref="XLPivotTable.ColumnLabels"/>.
/// </summary>
internal class XLPivotTableAxisField : IXLPivotField
{
    private readonly XLPivotTable _pivotTable;
    private readonly FieldIndex _index;

    internal XLPivotTableAxisField(XLPivotTable pivotTable, FieldIndex index)
    {
        _pivotTable = pivotTable;
        _index = index;
    }

    public string SourceName
    {
        get
        {
            if (_index.IsDataField)
                return XLConstants.PivotTable.ValuesSentinalLabel;

            return _pivotTable.PivotCache.FieldNames[_index];
        }
    }

    public string CustomName
    {
        get => GetFieldValue(f => f.Name, _pivotTable.DataCaption);
        set
        {
            if (_index.IsDataField)
            {
                _pivotTable.DataCaption = value;
                return;
            }

            if (_pivotTable.TryGetCustomNameFieldIndex(value, out var idx) && idx != _index)
                throw new ArgumentException($"Custom name '{value}' is already used by another field.");

            _pivotTable.PivotFields[_index].Name = value;
        }
    }

    public string SubtotalCaption
    {
        get => GetFieldValue(f => f.SubtotalCaption, string.Empty);
        set => GetField().SubtotalCaption = value;
    }

    public IReadOnlyCollection<XLSubtotalFunction> Subtotals
    {
        get
        {
            var subtotal = GetField().Subtotals;
            var isCustomSubtotal = subtotal.Count > 1 && subtotal.Contains(XLSubtotalFunction.Automatic);
            if (isCustomSubtotal)
            {
                // When subtotal is custom, the automatic is not shown
                subtotal = new HashSet<XLSubtotalFunction>(subtotal);
                subtotal.Remove(XLSubtotalFunction.Automatic);
            }

            return subtotal;
        }
    }

    public bool IncludeNewItemsInFilter
    {
        get => GetFieldValue(f => f.IncludeNewItemsInFilter, false);
        set => GetField().IncludeNewItemsInFilter = value;
    }

    public bool Outline
    {
        get => GetFieldValue(f => f.Outline, true);
        set => GetField().Outline = value;
    }
    public bool Compact
    {
        get => GetFieldValue(f => f.Compact, true);
        set => GetField().Compact = value;
    }

    public bool? SubtotalsAtTop
    {
        get => GetFieldValue(f => f.SubtotalTop, true);
        set => GetField().SubtotalTop = value ?? true;
    }

    public bool RepeatItemLabels
    {
        get => GetFieldValue(f => f.RepeatItemLabels, false);
        set => GetField().RepeatItemLabels = value;
    }

    public bool InsertBlankLines
    {
        get => GetFieldValue(f => f.InsertBlankRow, false);
        set => GetField().InsertBlankRow = value;
    }

    public bool ShowBlankItems
    {
        get => GetFieldValue(f => f.ShowAll, true);
        set => GetField().ShowAll = value;
    }

    public bool InsertPageBreaks
    {
        get => GetFieldValue(f => f.InsertPageBreak, false);
        set => GetField().InsertPageBreak = value;
    }

    public bool Collapsed
    {
        get
        {
            return GetFieldValue(f => !f.Items.Any(i => i.ShowDetails), false);
        }
        set
        {
            foreach (var item in GetField().Items)
                item.ShowDetails = !value;
        }
    }

    public XLPivotSortType SortType
    {
        get => GetFieldValue(f => f.SortType, XLPivotSortType.Default);
        set => GetField().SortType = value;
    }

    public IReadOnlyList<XLCellValue> SelectedValues => Array.Empty<XLCellValue>();

    public IXLPivotFieldStyleFormats StyleFormats => new XLPivotTableAxisFieldStyleFormats();

    public bool IsOnRowAxis => GetFieldValue(f => f.Axis == XLPivotAxis.AxisRow, _pivotTable.DataOnRows);

    public bool IsOnColumnAxis => GetFieldValue(f => f.Axis == XLPivotAxis.AxisCol, !_pivotTable.DataOnRows);

    public bool IsInFilterList => GetFieldValue(f => f.Axis == XLPivotAxis.AxisPage, false);

    public int Offset => _index;

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
        Collapsed = true;
        return this;
    }

    public IXLPivotField SetSort(XLPivotSortType value)
    {
        SortType = value;
        return this;
    }

    public IXLPivotField AddSelectedValue(XLCellValue value) => this;

    public IXLPivotField AddSelectedValues(IEnumerable<XLCellValue> values) => this;

    private XLPivotTableField GetField()
    {
        if (_index.IsDataField)
            throw new InvalidOperationException("Can't set this property on a data field.");

        return _pivotTable.PivotFields[_index];
    }

    private T GetFieldValue<T>(Func<XLPivotTableField, T> getter, T dataFieldValue)
    {
        if (_index.IsDataField)
            return dataFieldValue;
        var field = _pivotTable.PivotFields[_index];
        return getter(field);
    }
}
