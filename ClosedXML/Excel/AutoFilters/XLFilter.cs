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

        /// <summary>
        /// Value for <see cref="XLFilterType.Custom"/> that is compared using <see cref="Operator"/>.
        /// </summary>
        public XLCellValue CustomValue { get; set; }

        public Func<XLCellValue, bool> NewCondition { get; set; }

        public XLFilterOperator Operator { get; set; }

        public Object Value { get; set; }

        internal static XLFilter CreateCustomFilter(XLCellValue value, XLFilterOperator op, XLConnector connector)
        {
            // Keep in closure, so it doesn't have to be checked for every cell.
            var comparer = StringComparer.CurrentCultureIgnoreCase;
            return new XLFilter
            {
                CustomValue = value,
                Operator = op,
                Connector = connector,
                NewCondition = cellValue => CustomFilterSatisfied(cellValue, op, value, comparer),
            };
        }
        
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

        private static bool CustomFilterSatisfied(XLCellValue cellValue, XLFilterOperator op, XLCellValue filterValue, StringComparer textComparer)
        {
            // Blanks are rather strange case. Excel parsing logic for custom filter value into
            // XLCellValue is very inconsistent. E.g. 'does not equal' for empty string ignores
            // blanks and empty strings.
            cellValue = cellValue.IsBlank ? string.Empty : cellValue;
            filterValue = filterValue.IsBlank ? string.Empty : filterValue;

            if (cellValue.Type != filterValue.Type && cellValue.IsUnifiedNumber != filterValue.IsUnifiedNumber)
                return false;

            // Note that custom filter even error values, basically everything as a number.
            var comparison = cellValue.Type switch
            {
                XLDataType.Text => textComparer.Compare(cellValue.GetText(), filterValue.GetText()),
                XLDataType.Boolean => cellValue.GetBoolean().CompareTo(filterValue.GetBoolean()),
                XLDataType.Error => cellValue.GetError().CompareTo(filterValue.GetError()),
                _ => cellValue.GetUnifiedNumber().CompareTo(filterValue.GetUnifiedNumber())
            };

            // !!! Deviation from Excel !!!
            // Excel interprets custom filter with `equal` operator (and *only* equal operator) as
            // comparison of formatted string of a cell value with wildcard represented by custom
            // value filter value.
            // We do the sane thing and compare them for equality, so $10 is equal to 10.
            return op switch
            {
                XLFilterOperator.LessThan => comparison < 0,
                XLFilterOperator.EqualOrLessThan => comparison <= 0,
                XLFilterOperator.Equal => comparison == 0,
                XLFilterOperator.NotEqual => comparison != 0,
                XLFilterOperator.EqualOrGreaterThan => comparison >= 0,
                XLFilterOperator.GreaterThan => comparison > 0,
                _ => throw new NotSupportedException(),
            };
        }
    }
}
