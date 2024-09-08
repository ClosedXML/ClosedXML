// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;

namespace ClosedXML.Excel;

internal class XLPivotValueStyleFormat : XLPivotStyleFormatBase, IXLPivotValueStyleFormat
{
    /// <summary>
    /// A list of references that specify which data cells will be styled.
    /// A data cell will be styled, if it lies on all referenced fields.
    /// The term "lie on" means that either column or a row of data cell
    /// intersects a label cell of referenced field.
    /// </summary>
    private readonly List<FieldReference> _fieldReferences = new();

    public XLPivotValueStyleFormat(XLPivotTable pivotTable, FieldIndex fieldIndex)
        : base(pivotTable)
    {
        _fieldReferences.Add(new FieldReference(fieldIndex));
    }

    #region IXLPivotValueStyleFormat members

    public IXLPivotValueStyleFormat AndWith(IXLPivotField field)
    {
        _fieldReferences.Add(new FieldReference(field.Offset));
        return this;
    }

    public IXLPivotValueStyleFormat AndWith(IXLPivotField field, Predicate<XLCellValue>? predicate)
    {
        throw new NotImplementedException();
    }

    public IXLPivotValueStyleFormat ForValueField(IXLPivotValue valueField)
    {
        var valuesIndex = PivotTable.DataFields.IndexOf(valueField);
        if (valuesIndex == -1)
            throw new ArgumentOutOfRangeException($"Field '{valueField.CustomName}' is not among value fields of the pivot table.");

        _fieldReferences.Add(new FieldReference(FieldIndex.DataField, new[] { (uint)valuesIndex }));
        return this;
    }

    #endregion IXLPivotValueStyleFormat members

    internal override XLPivotArea GetCurrentArea()
    {
        var area = new XLPivotArea();
        foreach (var fieldReference in _fieldReferences)
        {
            var reference = new XLPivotReference
            {
                Field = unchecked((uint?)fieldReference.FieldIndex.Value)
            };
            if (fieldReference.Items is not null)
            {
                foreach (var item in fieldReference.Items)
                    reference.AddFieldItem(item);
            }

            area.AddReference(reference);
        }

        return area;
    }

    internal override bool Filter(XLPivotArea area)
    {
        var currentArea = GetCurrentArea();
        return XLPivotAreaComparer.Instance.Equals(area, currentArea);
    }

    private record FieldReference(FieldIndex FieldIndex, uint[]? Items = null);
}
