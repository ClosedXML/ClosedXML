// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    internal class XLPivotValueStyleFormat : XLPivotStyleFormat, IXLPivotValueStyleFormat
    {
        public XLPivotValueStyleFormat(IXLPivotField? field = null)
            : base(field)
        {
        }

        #region IXLPivotValueStyleFormat members

        public IXLPivotValueStyleFormat AndWith(IXLPivotField field)
        {
            return AndWith(field, null);
        }

        public IXLPivotValueStyleFormat AndWith(IXLPivotField field, Predicate<XLCellValue>? predicate)
        {
            FieldReferences.Add(new PivotLabelFieldReference(field, predicate));
            return this;
        }

        public IXLPivotValueStyleFormat ForValueField(IXLPivotValue valueField)
        {
            FieldReferences.Add(new PivotValueFieldReference(valueField.CustomName));
            return this;
        }

        #endregion IXLPivotValueStyleFormat members
    }
}
