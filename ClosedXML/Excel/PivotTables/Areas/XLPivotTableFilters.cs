#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections;
using System.Collections.Generic;

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
    /// <see cref="XLPivotTable.FilterFieldsPageWrap"/>
    /// </summary>
    private readonly List<FieldIndex> _fields = new();

    internal XLPivotTableFilters(XLPivotTable pivotTable)
    {
        _pivotTable = pivotTable;
    }

    IXLPivotField IXLPivotFields.Add(String sourceName) => Add(sourceName, sourceName);

    IXLPivotField IXLPivotFields.Add(String sourceName, String customName) => Add(sourceName, customName);

    public void Clear()
    {
        foreach (var fieldIndex in _fields)
            _pivotTable.RemoveFieldFromAxis(fieldIndex);

        _fields.Clear();
    }

    public Boolean Contains(String sourceName)
    {
        if (!_pivotTable.TryGetSourceNameFieldIndex(sourceName, out var index))
            return false;

        return _fields.Contains(index);
    }

    public bool Contains(IXLPivotField pivotField)
    {
        return Contains(pivotField.SourceName);
    }

    public IXLPivotField Get(String sourceName)
    {
        if (!_pivotTable.TryGetSourceNameFieldIndex(sourceName, out var index) ||
            !_fields.Contains(index))
            throw new KeyNotFoundException($"Field with source name '{sourceName}' not found in {XLPivotAxis.AxisPage}.");

        return new XLPivotTablePageField(_pivotTable, index);
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
        foreach (var fieldIndex in _fields)
            yield return new XLPivotTablePageField(_pivotTable, fieldIndex);
    }

    public Int32 IndexOf(String sourceName)
    {
        if (!_pivotTable.TryGetSourceNameFieldIndex(sourceName, out var fieldIndex))
            return -1;

        return _fields.IndexOf(fieldIndex);
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

        _fields.RemoveAt(index);
    }

    internal IReadOnlyList<FieldIndex> Fields => _fields;

    internal XLPivotTablePageField Add(String sourceName, String customName)
    {
        if (sourceName == XLConstants.PivotTable.ValuesSentinalLabel)
            throw new ArgumentException(nameof(sourceName), $"The column '{sourceName}' does not appear in the source range.");

        var index = _pivotTable.AddFieldToAxis(sourceName, customName, XLPivotAxis.AxisPage);
        _fields.Add(index);
        return new XLPivotTablePageField(_pivotTable, index);
    }
}
