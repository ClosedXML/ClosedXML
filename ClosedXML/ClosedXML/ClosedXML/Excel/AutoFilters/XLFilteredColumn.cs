using System;
using System.Linq;
namespace ClosedXML.Excel
{
    using System.Collections.Generic;

    internal class XLFilteredColumn: IXLFilteredColumn
    {
        XLAutoFilter _autoFilter;
        Int32 _column;
        public XLFilteredColumn(XLAutoFilter autoFilter, Int32 column)
        {
            _autoFilter = autoFilter;
            _column = column;
        }

        public IXLFilteredColumn AddFilter<T>(T value) where T: IComparable<T>
        {
            Func<Object, Boolean> condition;
            Boolean isText;
            if (typeof(T) == typeof(String))
            {
                condition = v => v.ToString().Equals(value.ToString(), StringComparison.InvariantCultureIgnoreCase);
                isText = true;
            }
            else
            {
                condition = v => (v.CastTo<T>() as IComparable).CompareTo(value) == 0;
                isText = false;
            }

            _autoFilter.Filters[_column].Add(new XLFilter { Value = value, Condition = condition, Operator = XLFilterOperator.Equal, Connector = XLConnector.Or });
            foreach (var row in _autoFilter.Range.Rows().Where(r => r.RowNumber() > 1))
            {
                if ((isText && condition(row.Cell(_column).GetString())) || (
                    !isText && row.Cell(_column).DataType == XLCellValues.Number && condition(row.Cell(_column).GetValue<T>()))
                    )
                    row.WorksheetRow().Unhide();
            }
            return this;
        }
    }
}