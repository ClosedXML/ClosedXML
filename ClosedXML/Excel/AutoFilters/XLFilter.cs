#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;
using ClosedXML.Excel.CalcEngine;

namespace ClosedXML.Excel
{
    internal enum XLConnector { And, Or }

    internal enum XLFilterOperator { Equal, NotEqual, GreaterThan, LessThan, EqualOrGreaterThan, EqualOrLessThan }

    internal class XLFilter
    {
        public XLConnector Connector { get; set; }

        public XLDateTimeGrouping DateTimeGrouping { get; set; }

        /// <summary>
        /// Value for <see cref="XLFilterType.Custom"/> that is compared using <see cref="Operator"/>.
        /// </summary>
        public XLCellValue CustomValue { get; init; }

        public Func<IXLCell, bool> Condition { get; set; }

        public XLFilterOperator Operator { get; init; } = XLFilterOperator.Equal;

        /// <summary>
        /// Value for <see cref="XLFilterType.Regular"/> filter.
        /// </summary>
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
                Condition = cell => CustomFilterSatisfied(cell.CachedValue, op, value, comparer),
            };
        }

        internal static XLFilter CreateWildcardFilter(string wildcard, bool match, XLConnector connector)
        {
            return new XLFilter
            {
                CustomValue = wildcard,
                Operator = match ? XLFilterOperator.Equal : XLFilterOperator.NotEqual,
                Connector = connector,
                Condition = match ? c => MatchesWildcard(wildcard, c.CachedValue) : c => !MatchesWildcard(wildcard, c.CachedValue),
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
                Condition = v => v.GetFormattedString().Equals(value.ToString(), StringComparison.OrdinalIgnoreCase), // TODO: Use cached value for formatted string.
            };
        }

        internal static XLFilter CreateDateGroupFilter(DateTime date, XLDateTimeGrouping dateTimeGrouping)
        {
            return new XLFilter
            {
                Value = date,
                Condition = cell => cell.CachedValue.IsDateTime && IsMatch(date, (DateTime)cell.CachedValue.GetDateTime(), dateTimeGrouping),
                Operator = XLFilterOperator.Equal,
                Connector = XLConnector.Or,
                DateTimeGrouping = dateTimeGrouping
            };
        }

        private static bool MatchesWildcard(string pattern, XLCellValue cellValue)
        {
            // Wildcard matches only text cells.
            if (!cellValue.IsText)
                return false;

            var text = cellValue.GetText();
            var wildcard = new Wildcard(pattern);
            var position = wildcard.Search(text.AsSpan());
            return position >= 0;
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

        internal static XLFilter CreateTopBottom(bool takeTop, double topBottomValue)
        {
            bool TopFilter(IXLCell cell)
            {
                var cachedValue = cell.CachedValue;
                return cachedValue.IsUnifiedNumber && cachedValue.GetUnifiedNumber() >= topBottomValue;
            }
            bool BottomFilter(IXLCell cell)
            {
                var cachedValue = cell.CachedValue;
                return cachedValue.IsUnifiedNumber && cachedValue.GetUnifiedNumber() <= topBottomValue;
            }

            return new XLFilter
            {
                Value = topBottomValue,
                Operator = XLFilterOperator.Equal,
                Connector = XLConnector.Or,
                Condition = takeTop ? TopFilter : BottomFilter,
            };
        }
    }
}
