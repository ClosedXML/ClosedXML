// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    internal class XLPivotValueStyleFormat : XLPivotStyleFormat, IXLPivotValueStyleFormat
    {
        public XLPivotValueStyleFormat(IXLPivotField field = null, IXLStyle style = null)
            : base(field, style)
        { }

        #region IXLPivotValueStyleFormat members

        public IXLPivotValueStyleFormat AndWith(IXLPivotField field)
        {
            return AndWith(field, null);
        }

        public IXLPivotValueStyleFormat AndWith(IXLPivotField field, Predicate<Object> predicate)
        {
            FieldReferences.Add(new PivotLabelFieldReference(field, predicate));
            return this;
        }

        public IXLPivotValueStyleFormat ForValueField(IXLPivotValue valueField)
        {
            FieldReferences.Add(new PivotValueFieldReference(valueField.SourceName));
            return this;
        }

        #endregion IXLPivotValueStyleFormat members
    }
}
