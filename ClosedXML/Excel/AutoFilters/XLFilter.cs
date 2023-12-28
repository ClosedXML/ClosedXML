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

        internal static XLFilter CreateRegularDateGroupFilter(DateTime date, XLDateTimeGrouping dateTimeGrouping)
        {
            return new XLFilter
            {
                Value = date,
                Condition = date2 => IsMatch(date, (DateTime)date2, dateTimeGrouping),
                Operator = XLFilterOperator.Equal,
                Connector = XLConnector.Or,
                DateTimeGrouping = dateTimeGrouping
            };
        }

        private static Boolean IsMatch(DateTime date1, DateTime date2, XLDateTimeGrouping dateTimeGrouping)
        {
            Boolean isMatch = true;
            if (isMatch && dateTimeGrouping >= XLDateTimeGrouping.Year) isMatch &= date1.Year.Equals(date2.Year);
            if (isMatch && dateTimeGrouping >= XLDateTimeGrouping.Month) isMatch &= date1.Month.Equals(date2.Month);
            if (isMatch && dateTimeGrouping >= XLDateTimeGrouping.Day) isMatch &= date1.Day.Equals(date2.Day);
            if (isMatch && dateTimeGrouping >= XLDateTimeGrouping.Hour) isMatch &= date1.Hour.Equals(date2.Hour);
            if (isMatch && dateTimeGrouping >= XLDateTimeGrouping.Minute) isMatch &= date1.Minute.Equals(date2.Minute);
            if (isMatch && dateTimeGrouping >= XLDateTimeGrouping.Second) isMatch &= date1.Second.Equals(date2.Second);

            return isMatch;
        }
    }
}
