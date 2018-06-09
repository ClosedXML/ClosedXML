// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    internal enum XLConnector { And, Or }

    internal enum XLFilterOperator { Equal, NotEqual, GreaterThan, LessThan, EqualOrGreaterThan, EqualOrLessThan }

    internal class XLFilter
    {
        public XLFilter(XLFilterOperator op = XLFilterOperator.Equal)
        {
            Operator = op;
        }

        public Func<Object, Boolean> Condition { get; set; }
        public XLConnector Connector { get; set; }
        public XLDateTimeGrouping DateTimeGrouping { get; set; }
        public XLFilterOperator Operator { get; set; }
        public Object Value { get; set; }
    }
}
