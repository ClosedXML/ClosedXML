using System;

namespace ClosedXML.Excel
{
    public interface IXLDateTimeGroupFilteredColumn
    {
        IXLDateTimeGroupFilteredColumn AddDateGroupFilter(DateTime date, XLDateTimeGrouping dateTimeGrouping);
    }

    internal class XLDateTimeGroupFilteredColumn : IXLDateTimeGroupFilteredColumn
    {
        private readonly XLAutoFilter _autoFilter;
        private readonly int _column;

        public XLDateTimeGroupFilteredColumn(XLAutoFilter autoFilter, int column)
        {
            _autoFilter = autoFilter;
            _column = column;
        }

        public IXLDateTimeGroupFilteredColumn AddDateGroupFilter(DateTime date, XLDateTimeGrouping dateTimeGrouping)
        {
            bool condition(object date2) => IsMatch(date, (DateTime)date2, dateTimeGrouping);

            _autoFilter.Filters[_column].Add(new XLFilter
            {
                Value = date,
                Condition = condition,
                Operator = XLFilterOperator.Equal,
                Connector = XLConnector.Or,
                DateTimeGrouping = dateTimeGrouping
            });

            var rows = _autoFilter.Range.Rows(2, _autoFilter.Range.RowCount());
            foreach (var row in rows)
            {
                if (row.Cell(_column).DataType == XLDataType.DateTime && condition(row.Cell(_column).GetDateTime()))
                {
                    row.WorksheetRow().Unhide();
                }
            }

            return this;
        }

        internal static bool IsMatch(DateTime date1, DateTime date2, XLDateTimeGrouping dateTimeGrouping)
        {
            var isMatch = true;
            if (isMatch && dateTimeGrouping >= XLDateTimeGrouping.Year)
            {
                isMatch &= date1.Year.Equals(date2.Year);
            }

            if (isMatch && dateTimeGrouping >= XLDateTimeGrouping.Month)
            {
                isMatch &= date1.Month.Equals(date2.Month);
            }

            if (isMatch && dateTimeGrouping >= XLDateTimeGrouping.Day)
            {
                isMatch &= date1.Day.Equals(date2.Day);
            }

            if (isMatch && dateTimeGrouping >= XLDateTimeGrouping.Hour)
            {
                isMatch &= date1.Hour.Equals(date2.Hour);
            }

            if (isMatch && dateTimeGrouping >= XLDateTimeGrouping.Minute)
            {
                isMatch &= date1.Minute.Equals(date2.Minute);
            }

            if (isMatch && dateTimeGrouping >= XLDateTimeGrouping.Second)
            {
                isMatch &= date1.Second.Equals(date2.Second);
            }

            return isMatch;
        }
    }
}