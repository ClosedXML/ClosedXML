#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;
using System.Globalization;
using ClosedXML.Excel.CalcEngine;

namespace ClosedXML.Excel
{
    internal enum XLConnector { And, Or }

    internal enum XLFilterOperator { Equal, NotEqual, GreaterThan, LessThan, EqualOrGreaterThan, EqualOrLessThan }

    /// <summary>
    /// A single filter condition for auto filter.
    /// </summary>
    internal class XLFilter
    {
        private XLFilter()
        {
        }

        public XLConnector Connector { get; set; }

        public XLDateTimeGrouping DateTimeGrouping { get; set; }

        /// <summary>
        /// Value for <see cref="XLFilterType.Custom"/> that is compared using <see cref="Operator"/>.
        /// </summary>
        public XLCellValue CustomValue { get; init; }

        public Func<IXLCell, XLFilterColumn, bool> Condition { get; init; }

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
                Condition = (cell, _) => CustomFilterSatisfied(cell.CachedValue, op, value, comparer),
            };
        }

        internal static XLFilter CreateCustomPatternFilter(string filterValue, bool match, XLConnector connector)
        {
            // Excel really parses value in current culture to detect type (e.g. 1,00 is detected as a number in cs-CZ).
            var testValue = XLCellValue.FromText(filterValue, CultureInfo.CurrentCulture);
            if (testValue.Type == XLDataType.Text)
            {
                // Custom filter Equal matches strings with a pattern. Custom uses it mostly for filters like begin-with (e.g. `ABC*`).
                var wildcard = filterValue;
                return new XLFilter
                {
                    CustomValue = wildcard,
                    Operator = match ? XLFilterOperator.Equal : XLFilterOperator.NotEqual,
                    Connector = connector,
                    Condition = match ? (c, _) => TextMatchesWildcard(wildcard, c) : (c, _) => !TextMatchesWildcard(wildcard, c),
                };
            }

            // Keep in closure, so it doesn't have to be checked for every cell.
            var comparer = StringComparer.CurrentCultureIgnoreCase;
            return new XLFilter
            {
                CustomValue = filterValue,
                Operator = match ? XLFilterOperator.Equal : XLFilterOperator.NotEqual,
                Connector = connector,
                Condition = match ? (c, _) => ContentMatches(c, filterValue)
                    : (c, _) => !CustomFilterSatisfied(c.CachedValue, XLFilterOperator.Equal, testValue, comparer),
            };

            static bool TextMatchesWildcard(string pattern, IXLCell cell)
            {
                var cachedValue = cell.CachedValue;
                if (!cachedValue.IsText)
                    return false;

                var wildcard = new Wildcard(pattern);
                var position = wildcard.Search(cachedValue.GetText().AsSpan());
                return position >= 0;
            }
        }

        internal static XLFilter CreateRegularFilter(string filterValue)
        {

            return new XLFilter
            {
                Value = filterValue,
                Operator = XLFilterOperator.Equal,
                Connector = XLConnector.Or,
                Condition = (cell, _) => ContentMatches(cell, filterValue)
            };
        }

        internal static XLFilter CreateDateGroupFilter(DateTime date, XLDateTimeGrouping dateTimeGrouping)
        {
            return new XLFilter
            {
                Value = date,
                Condition = HasSameGroup,
                Operator = XLFilterOperator.Equal,
                Connector = XLConnector.Or,
                DateTimeGrouping = dateTimeGrouping
            };

            bool HasSameGroup(IXLCell cell, XLFilterColumn _)
            {
                var cachedValue = cell.CachedValue;
                return cachedValue.IsDateTime && IsMatch(date, cachedValue.GetDateTime(), dateTimeGrouping);
            }

            static Boolean IsMatch(DateTime date1, DateTime date2, XLDateTimeGrouping dateTimeGrouping)
            {
                Boolean isMatch = true;
                if (dateTimeGrouping >= XLDateTimeGrouping.Year) isMatch &= date1.Year.Equals(date2.Year);
                if (isMatch && dateTimeGrouping >= XLDateTimeGrouping.Month) isMatch &= date1.Month.Equals(date2.Month);
                if (isMatch && dateTimeGrouping >= XLDateTimeGrouping.Day) isMatch &= date1.Day.Equals(date2.Day);
                if (isMatch && dateTimeGrouping >= XLDateTimeGrouping.Hour) isMatch &= date1.Hour.Equals(date2.Hour);
                if (isMatch && dateTimeGrouping >= XLDateTimeGrouping.Minute) isMatch &= date1.Minute.Equals(date2.Minute);
                if (isMatch && dateTimeGrouping >= XLDateTimeGrouping.Second) isMatch &= date1.Second.Equals(date2.Second);

                return isMatch;
            }
        }

        private static bool ContentMatches(IXLCell cell, string filterValue)
        {
            // IXLCell.GetFormattedString() could trigger formula evaluation.
            var cachedValue = cell.CachedValue;
            var formattedString = ((XLCell)cell).GetFormattedString(cachedValue);
            return formattedString.Equals(filterValue, StringComparison.OrdinalIgnoreCase);
        }

        private static bool CustomFilterSatisfied(XLCellValue cellValue, XLFilterOperator op, XLCellValue filterValue, StringComparer textComparer)
        {
            // Blanks are rather strange case. Excel parsing logic for custom filter value into
            // XLCellValue is very inconsistent. E.g. 'does not equal' for empty string ignores
            // blanks and empty strings.
            // For custom compare filters, blank never matches.
            if (cellValue.IsBlank || filterValue.IsBlank)
                return false;

            if (cellValue.Type != filterValue.Type)
            {
                // Types are different, but could still be unified numbers and thus comparable.
                if (!(cellValue.IsUnifiedNumber && filterValue.IsUnifiedNumber))
                    return false;
            }

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

        internal static XLFilter CreateTopBottom(bool takeTop, int percentsOrItemCount)
        {
            bool TopFilter(IXLCell cell, XLFilterColumn filterColumn)
            {
                var cachedValue = cell.CachedValue;
                return cachedValue.IsUnifiedNumber && cachedValue.GetUnifiedNumber() >= filterColumn.TopBottomFilterValue;
            }
            bool BottomFilter(IXLCell cell, XLFilterColumn filterColumn)
            {
                var cachedValue = cell.CachedValue;
                return cachedValue.IsUnifiedNumber && cachedValue.GetUnifiedNumber() <= filterColumn.TopBottomFilterValue;
            }

            return new XLFilter
            {
                Value = percentsOrItemCount,
                Operator = XLFilterOperator.Equal,
                Connector = XLConnector.Or,
                Condition = takeTop ? TopFilter : BottomFilter,
            };
        }

        internal static XLFilter CreateAverage(double initialAverage, bool aboveAverage)
        {
            bool AboveAverage(IXLCell cell, XLFilterColumn filterColumn)
            {
                var cachedValue = cell.CachedValue;
                var average = filterColumn.DynamicValue;
                return cachedValue.IsUnifiedNumber && cachedValue.GetUnifiedNumber() > average;
            }
            bool BelowAverage(IXLCell cell, XLFilterColumn filterColumn)
            {
                var cachedValue = cell.CachedValue;
                var average = filterColumn.DynamicValue;
                return cachedValue.IsUnifiedNumber && cachedValue.GetUnifiedNumber() < average;
            }

            return new XLFilter
            {
                Value = initialAverage,
                Operator = XLFilterOperator.Equal,
                Connector = XLConnector.Or,
                Condition = aboveAverage ? AboveAverage : BelowAverage,
            };
        }
    }
}
