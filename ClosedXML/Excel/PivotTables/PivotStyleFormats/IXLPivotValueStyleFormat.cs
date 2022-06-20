// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    public interface IXLPivotValueStyleFormat : IXLPivotStyleFormat
    {
        IXLPivotValueStyleFormat AndWith(IXLPivotField field);

        IXLPivotValueStyleFormat AndWith(IXLPivotField field, Predicate<object> predicate);

        IXLPivotValueStyleFormat ForValueField(IXLPivotValue valueField);
    }
}
