// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace ClosedXML.Excel;

/// <summary>
/// Page/filter fields of a <see cref="XLPivotTable"/>. It determines filter values and layout.
/// It is accessible through fluent API <see cref="XLPivotTable.ReportFilters"/>.
/// </summary>
internal class XLPivotTableFilters : IXLPivotFields
{
    private readonly XLPivotTable _pivotTable;

    /// <summary>
    /// Filter fields in correct order. The layout is determined by
    /// <see cref="XLPivotTable.FilterFieldsPageWrap"/> and
    /// <see cref="XLPivotTable.FilterAreaOrder"/>.
    /// </summary>
    private readonly List<XLPivotPageField> _fields = new();

    internal XLPivotTableFilters(XLPivotTable pivotTable)
    {
        _pivotTable = pivotTable;
    }

    IXLPivotField IXLPivotFields.Add(String sourceName) => Add(sourceName, sourceName);

    IXLPivotField IXLPivotFields.Add(String sourceName, String customName) => Add(sourceName, customName);

    public void Clear()
    {
        foreach (var field in _fields)
            _pivotTable.RemoveFieldFromAxis(field.Field);

        _fields.Clear();
    }

    public Boolean Contains(String sourceName)
    {
        return IndexOf(sourceName) >= 0;
    }

    public bool Contains(IXLPivotField pivotField)
    {
        return Contains(pivotField.SourceName);
    }

    public IXLPivotField Get(String sourceName)
    {
        if (!_pivotTable.TryGetSourceNameFieldIndex(sourceName, out var fieldIndex))
            throw new KeyNotFoundException($"Field with source name '{sourceName}' not found in {XLPivotAxis.AxisPage}.");

        var filterField = _fields.SingleOrDefault(f => f.Field == fieldIndex);
        if (filterField is null)
            throw new KeyNotFoundException($"Field with source name '{sourceName}' not found in {XLPivotAxis.AxisPage}.");

        return new XLPivotTablePageField(_pivotTable, filterField);
    }

    public IXLPivotField Get(Int32 index)
    {
        if (index < 0 || index >= _fields.Count)
            throw new IndexOutOfRangeException();

        return new XLPivotTablePageField(_pivotTable, _fields[index]);
    }

    IEnumerator<IXLPivotField> IEnumerable<IXLPivotField>.GetEnumerator() => GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    public IEnumerator<XLPivotTablePageField> GetEnumerator()
    {
        foreach (var field in _fields)
            yield return new XLPivotTablePageField(_pivotTable, field);
    }

    public Int32 IndexOf(String sourceName)
    {
        if (!_pivotTable.TryGetSourceNameFieldIndex(sourceName, out var fieldIndex))
            return -1;

        return _fields.FindIndex(f => f.Field == fieldIndex);
    }

    public Int32 IndexOf(IXLPivotField pf)
    {
        return IndexOf(pf.SourceName);
    }

    public void Remove(String sourceName)
    {
        var index = IndexOf(sourceName);
        if (index == -1)
            return;

        var removedRows = _fields.Count > 1 ? 1 : 2;
        var movedArea = _pivotTable.Area.ShiftRows(-removedRows);

        _fields.RemoveAt(index);
        _pivotTable.RemoveFieldFromAxis(index);

        _pivotTable.Area = movedArea;
    }

    internal IReadOnlyList<XLPivotPageField> Fields => _fields;

    internal XLPivotTablePageField Add(String sourceName, String customName)
    {
        if (sourceName == XLConstants.PivotTable.ValuesSentinalLabel)
            throw new ArgumentException(nameof(sourceName), $"The column '{sourceName}' does not appear in the source range.");

        var addedRows = _fields.Count > 0 ? 1 : 2;
        var movedArea = _pivotTable.Area.ShiftRows(addedRows);
        
        var fieldIndex = _pivotTable.AddFieldToAxis(sourceName, customName, XLPivotAxis.AxisPage);
        var filterField = new XLPivotPageField(fieldIndex);
        _fields.Add(filterField);

        _pivotTable.Area = movedArea;
        return new XLPivotTablePageField(_pivotTable, filterField);
    }

    internal bool Contains(FieldIndex fieldIndex)
    {
        return _fields.FindIndex(f => f.Field == fieldIndex) >= 0;
    }

    internal void AddField(XLPivotPageField pageField)
    {
        _fields.Add(pageField);
    }

    /// <summary>
    /// Number of rows/cols occupied by the filter area. Filter area is above the pivot table and it
    /// optional (i.e. size <c>0</c> indicates no filter).
    /// </summary>
    internal (int Width, int Height) GetSize()
    {
        var pageWrap = _pivotTable.FilterFieldsPageWrap;
        if (pageWrap == 0)
            pageWrap = int.MaxValue;

        var dim1 = Math.DivRem(_fields.Count, pageWrap, out var dim2);
        dim1 = _fields.Count > 0 ? dim1 + 1 : dim1;

        return _pivotTable.FilterAreaOrder switch
        {
            XLFilterAreaOrder.DownThenOver => new(dim1, dim2),
            XLFilterAreaOrder.OverThenDown => new(dim2, dim1),
            _ => throw new UnreachableException(),
        };
    }

    /// <summary>
    /// Number of rows/cols occupied by the filter area, including the gap below, if there is at least one filter.
    /// </summary>
    internal (int Width, int Height) GetSizeWithGap()
    {
        var filtersSize = GetSize();
        return filtersSize.Height > 0
            ? (filtersSize.Width, filtersSize.Height + 1)
            : filtersSize;
    }
}
