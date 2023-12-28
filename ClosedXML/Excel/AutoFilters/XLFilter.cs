#nullable disable

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

        internal static XLFilter CreateRegularFilter(XLCellValue value)
        {
            // TODO: If user supplies a text that is a wildcard, escape it (e.g. `2*` to `2~*`).
            var wildcard = value.ToString();
            return new XLFilter
            {
                Value = wildcard,
                Operator = XLFilterOperator.Equal,
                Connector = XLConnector.Or,
                Condition = v => v.ToString().Equals(value.ToString(), StringComparison.OrdinalIgnoreCase)
            };
        }
    }
}
