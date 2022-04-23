using System;

namespace ClosedXML.Excel
{
    internal class XLFilteredColumn : IXLFilteredColumn
    {
        private readonly XLAutoFilter _autoFilter;
        private readonly int _column;

        public XLFilteredColumn(XLAutoFilter autoFilter, int column)
        {
            _autoFilter = autoFilter;
            _column = column;
        }

        #region IXLFilteredColumn Members

        public IXLFilteredColumn AddFilter<T>(T value) where T : IComparable<T>
        {
            Func<object, bool> condition;
            bool isText;
            if (typeof(T) == typeof(string))
            {
                condition = v => v.ToString().Equals(value.ToString(), StringComparison.InvariantCultureIgnoreCase);
                isText = true;
            }
            else
            {
                condition = v => v.CastTo<T>().CompareTo(value) == 0;
                isText = false;
            }

            _autoFilter.Filters[_column].Add(new XLFilter
            {
                Value = value,
                Condition = condition,
                Operator = XLFilterOperator.Equal,
                Connector = XLConnector.Or
            });

            var rows = _autoFilter.Range.Rows(2, _autoFilter.Range.RowCount());

            foreach (var row in rows)
            {
                if ((isText && condition(row.Cell(_column).GetString())) ||
                    (!isText && row.Cell(_column).DataType == XLDataType.Number &&
                     condition(row.Cell(_column).GetValue<T>())))
                {
                    row.WorksheetRow().Unhide();
                }
            }
            return this;
        }

        #endregion IXLFilteredColumn Members
    }
}
