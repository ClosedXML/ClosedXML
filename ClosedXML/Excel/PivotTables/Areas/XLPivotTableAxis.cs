using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel;

/// <summary>
/// A description of one axis (<see cref="XLPivotTable.RowAxis"/>/<see cref="XLPivotTable.ColumnAxis"/>)
/// of a <see cref="XLPivotTable"/>. It consists of fields in a specific order and values that make up
/// individual rows/columns of the axis.
/// </summary>
/// <remarks>
/// [ISO-29500] 18.10.1.17 colItems (Column Items), 18.10.1.84 rowItems (Row Items).
/// </remarks>
internal class XLPivotTableAxis : IXLPivotFields
{
    private readonly XLPivotTable _pivotTable;

    private readonly XLPivotAxis _axis;

    /// <summary>
    /// Fields displayed on the axis, in the order of the fields on the axis.
    /// </summary>
    private readonly List<FieldIndex> _fields = new();

    /// <summary>
    /// Values of one row/column in an axis. Items are not kept in sync with <see cref="_fields"/>.
    /// </summary>
    private readonly List<XLPivotFieldAxisItem> _axisItems = new();

    internal XLPivotTableAxis(XLPivotTable pivotTable, XLPivotAxis axis)
    {
        _pivotTable = pivotTable;
        _axis = axis;
    }

    /// <summary>
    /// A list of fields to displayed on the axis. It determines which fields and in what order
    /// should the fields be displayed.
    /// </summary>
    internal IReadOnlyList<FieldIndex> Fields => _fields;

    /// <summary>
    /// Individual row/column parts of the axis.
    /// </summary>
    internal IReadOnlyList<XLPivotFieldAxisItem> Items => _axisItems;

    internal bool ContainsDataField => _fields.Any(x => x.IsDataField);

    IXLPivotField IXLPivotFields.Add(String sourceName) => Add(sourceName, sourceName);

    IXLPivotField IXLPivotFields.Add(String sourceName, String customName) => Add(sourceName, customName);

    void IXLPivotFields.Clear() => Clear();

    Boolean IXLPivotFields.Contains(String sourceName) => Contains(sourceName);

    Boolean IXLPivotFields.Contains(IXLPivotField pivotField) => Contains(pivotField.SourceName);

    IXLPivotField IXLPivotFields.Get(String sourceName)
    {
        if (!_pivotTable.TryGetSourceNameFieldIndex(sourceName, out var index) ||
            !_fields.Contains(index))
            throw new KeyNotFoundException($"Field with source name '{sourceName}' not found in {_axis}.");

        return new XLPivotTableAxisField(_pivotTable, index);
    }

    IXLPivotField IXLPivotFields.Get(Int32 index)
    {
        if (index < 0 || index >= _fields.Count)
            throw new IndexOutOfRangeException();

        return new XLPivotTableAxisField(_pivotTable, _fields[index]);
    }

    Int32 IXLPivotFields.IndexOf(String sourceName)
    {
        return IndexOf(sourceName);
    }

    Int32 IXLPivotFields.IndexOf(IXLPivotField pf)
    {
        return IndexOf(pf.SourceName);
    }

    void IXLPivotFields.Remove(String sourceName)
    {
        var index = IndexOf(sourceName);
        if (index == -1)
            return;

        _pivotTable.RemoveFieldFromAxis(_fields[index]);
        _fields.RemoveAt(index);
    }

    IEnumerator<IXLPivotField> IEnumerable<IXLPivotField>.GetEnumerator() => GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    public IEnumerator<XLPivotTableAxisField> GetEnumerator()
    {
        foreach (var fieldIndex in _fields)
            yield return new XLPivotTableAxisField(_pivotTable, fieldIndex);
    }

    internal int IndexOf(FieldIndex index)
    {
        return _fields.IndexOf(index);
    }

    internal bool Contains(string sourceName)
    {
        if (!_pivotTable.TryGetSourceNameFieldIndex(sourceName, out var index))
            return false;

        return _fields.Contains(index);
    }

    /// <summary>
    /// Add field to the axis, as an index.
    /// </summary>
    internal void AddField(FieldIndex fieldIndex)
    {
        if (_pivotTable.IsFieldUsedOnAxis(fieldIndex))
            throw new ArgumentException("Field is already used on an axis.");

        _fields.Add(fieldIndex);
    }

    private XLPivotTableAxisField Add(String sourceName, String customName)
    {
        var field = AddField(sourceName, customName);

        // Excel by default adds a subtotal, but previous versions of ClosedXML didn't have them,
        // so keep API behavior.
        if (field.Offset != FieldIndex.DataField.Value)
            _pivotTable.PivotFields[field.Offset].RemoveSubtotal(XLSubtotalFunction.Automatic);

        return field;
    }

    internal XLPivotTableAxisField AddField(String sourceName, String customName)
    {
        var index = _pivotTable.AddFieldToAxis(sourceName, customName, _axis);
        _fields.Add(index);
        return new XLPivotTableAxisField(_pivotTable, index);
    }

    /// <summary>
    /// Add a row/column axis values (i.e. values visible on the axis).
    /// </summary>
    internal void AddItem(XLPivotFieldAxisItem axisItem)
    {
        _axisItems.Add(axisItem);
    }

    internal void Clear()
    {
        foreach (var fieldIndex in _fields)
            _pivotTable.RemoveFieldFromAxis(fieldIndex);

        _axisItems.Clear();
        _fields.Clear();
    }

    private Int32 IndexOf(String sourceName)
    {
        if (!_pivotTable.TryGetSourceNameFieldIndex(sourceName, out var fieldIndex))
            return -1;

        return _fields.IndexOf(fieldIndex);
    }
}
