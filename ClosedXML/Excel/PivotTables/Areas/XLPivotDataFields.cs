using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel;

/// <summary>
/// A collection of <see cref="XLPivotDataField"/>.
/// </summary>
internal class XLPivotDataFields : IXLPivotValues, IReadOnlyCollection<XLPivotDataField>
{
    private readonly XLPivotTable _pivotTable;

    /// <summary>
    /// Fields displayed in the data area of the pivot table, in the order fields are displayed.
    /// </summary>
    private readonly List<XLPivotDataField> _fields = new();

    internal XLPivotDataFields(XLPivotTable pivotTable)
    {
        _pivotTable = pivotTable;
    }

    public int Count => _fields.Count;

    #region IXLPivotValues

    public IXLPivotValue Add(string sourceName)
    {
        return AddField(sourceName, sourceName);
    }

    public IXLPivotValue Add(string sourceName, string customName)
    {
        return AddField(sourceName, customName);
    }

    public void Clear()
    {
        _fields.Clear();
        foreach (var field in _fields)
            _pivotTable.RemoveFieldFromAxis(field.Field);
    }

    public bool Contains(string customName)
    {
        return IndexOf(customName) != -1;
    }

    public bool Contains(IXLPivotValue pivotValue)
    {
        return Contains(pivotValue.CustomName);
    }

    public IXLPivotValue Get(string customName)
    {
        var dataField = _fields.SingleOrDefault(x => XLHelper.NameComparer.Equals(x.CustomName, customName));
        if (dataField is null)
        {
            throw new KeyNotFoundException($"Unable to find data field for '{customName}'.");
        }

        return dataField;
    }

    public IXLPivotValue Get(int index)
    {
        return _fields[index];
    }

    public int IndexOf(string customName)
    {
        return _fields.FindIndex(x => XLHelper.NameComparer.Equals(x.CustomName, customName));
    }

    public int IndexOf(IXLPivotValue pivotValue)
    {
        return IndexOf(pivotValue.CustomName);
    }

    public void Remove(string customName)
    {
        var index = IndexOf(customName);
        if (index == -1)
            return;

        var dataField = _fields[index];
        _pivotTable.RemoveFieldFromAxis(dataField.Field);
        _fields.Remove(dataField);
    }

    IEnumerator<IXLPivotValue> IEnumerable<IXLPivotValue>.GetEnumerator()
    {
        return GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion

    internal XLPivotDataField AddField(string sourceName, string? customName)
    {
        if (!_pivotTable.TryGetSourceNameFieldIndex(sourceName, out var fieldIndex))
            throw new ArgumentOutOfRangeException($"Field '{sourceName}' is not in the pivot cache.");

        if (fieldIndex.IsDataField)
            throw new ArgumentException("'Values' field can be used only on row or column axis.");

        var dataField = new XLPivotDataField(_pivotTable, fieldIndex.Value)
        {
            DataFieldName = customName,
        };
        AddField(dataField);

        // If there are multiple values, at least axis must contain 'data' field.
        // Otherwise, Excel requires a repair.
        if (_fields.Count > 1 &&
            !_pivotTable.RowAxis.ContainsDataField &&
            !_pivotTable.ColumnAxis.ContainsDataField)
        {
            _pivotTable.ColumnLabels.Add(XLConstants.PivotTable.ValuesSentinalLabel);
        }

        return dataField;
    }

    internal void AddField(XLPivotDataField dataField)
    {
        // Excel invariant - data field must have the flag if and only if it is in the data fields collection.
        _fields.Add(dataField);
        _pivotTable.PivotFields[dataField.Field].DataField = true;
    }

    public IEnumerator<XLPivotDataField> GetEnumerator()
    {
        return _fields.GetEnumerator();
    }
}
