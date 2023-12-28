using System;

namespace ClosedXML.Excel
{
    internal class XLDateTimeGroupFilteredColumn : IXLDateTimeGroupFilteredColumn
    {
        private readonly XLAutoFilter _autoFilter;
        private readonly Int32 _column;

        public XLDateTimeGroupFilteredColumn(XLAutoFilter autoFilter, Int32 column)
        {
            _autoFilter = autoFilter;
            _column = column;
        }

        public IXLDateTimeGroupFilteredColumn AddDateGroupFilter(DateTime date, XLDateTimeGrouping dateTimeGrouping)
        {
            _autoFilter.AddFilter(_column, XLFilter.CreateRegularDateGroupFilter(date, dateTimeGrouping));
            _autoFilter.Reapply();
            return this;
        }
    }
}
