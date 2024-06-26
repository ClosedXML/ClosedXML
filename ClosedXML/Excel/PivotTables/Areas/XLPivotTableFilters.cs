#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections;
using System.Collections.Generic;
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
    /// <see cref="XLPivotTable.FilterFieldsPageWrap"/>
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
        if (!_pivotTable.TryGetSourceNameFieldIndex(sourceName, out var index))
            throw new KeyNotFoundException($"Field with source name '{sourceName}' not found in {XLPivotAxis.AxisPage}.");

        if (_fields.All(f => f.Field != index))
            throw new KeyNotFoundException($"Field with source name '{sourceName}' not found in {XLPivotAxis.AxisPage}.");

        return new XLPivotTablePageField(_pivotTable, index);
    }

    public IXLPivotField Get(Int32 index)
    {
        if (index < 0 || index >= _fields.Count)
            throw new IndexOutOfRangeException();

        return new XLPivotTablePageField(_pivotTable, _fields[index].Field);
    }

    IEnumerator<IXLPivotField> IEnumerable<IXLPivotField>.GetEnumerator() => GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    public IEnumerator<XLPivotTablePageField> GetEnumerator()
    {
        foreach (var field in _fields)
            yield return new XLPivotTablePageField(_pivotTable, field.Field);
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

        _fields.RemoveAt(index);
    }

    internal IReadOnlyList<XLPivotPageField> Fields => _fields;

    internal XLPivotTablePageField Add(String sourceName, String customName)
    {
        if (sourceName == XLConstants.PivotTable.ValuesSentinalLabel)
            throw new ArgumentException(nameof(sourceName), $"The column '{sourceName}' does not appear in the source range.");

        var index = _pivotTable.AddFieldToAxis(sourceName, customName, XLPivotAxis.AxisPage);
        _fields.Add(new XLPivotPageField(index));
        return new XLPivotTablePageField(_pivotTable, index);
    }

    internal bool Contains(FieldIndex fieldIndex)
    {
        return _fields.FindIndex(f => f.Field == fieldIndex) >= 0;
    }

    internal void AddField(XLPivotPageField pageField)
    {
        _fields.Add(pageField);
    }
}
