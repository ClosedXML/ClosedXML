using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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

        public XLFilterOperator Operator { get; set; }
        public Object Value { get; set; }
        public XLConnector Connector { get; set; }
        public Func<Object, Boolean> Condition { get; set; }
        public XLDateTimeGrouping DateTimeGrouping { get; set; }
    }
}
