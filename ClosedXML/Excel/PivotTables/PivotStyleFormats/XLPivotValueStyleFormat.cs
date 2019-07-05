// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal class XLPivotValueStyleFormat : IXLPivotValueStyleFormat
    {
        public XLPivotValueStyleFormat(IXLPivotFormat pivotFormat)
        {
            PivotFormat = pivotFormat;
        }

        #region IXLPivotValueStyleFormat members

        public IXLStyle Style
        {
            get => PivotFormat.Style;
            set => PivotFormat.Style = value;
        }

        public IXLPivotFormat PivotFormat { get; }

        public IXLPivotValueStyleFormat AndWith(IXLPivotField field)
        {
            return AndWith(field, null);
        }

        public IXLPivotValueStyleFormat AndWith(IXLPivotField field, Predicate<Object> predicate)
        {
            ((List<IFieldRef>)PivotFormat.FieldReferences).Add(FieldRef.ForField(field.SourceName, predicate));
            return this;
        }

        public IXLPivotValueStyleFormat ForValueField(IXLPivotValue valueField)
        {
            ((List<IFieldRef>)PivotFormat.FieldReferences).Add(FieldRef.ValueField(valueField.SourceName));
            return this;
        }

        #endregion IXLPivotValueStyleFormat members
    }
}
