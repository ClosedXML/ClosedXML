using System;

namespace ClosedXML.Excel
{
    internal class XLFilteredColumn : IXLFilteredColumn
    {
        private readonly XLAutoFilter _autoFilter;
        private readonly Int32 _column;

        public XLFilteredColumn(XLAutoFilter autoFilter, Int32 column)
        {
            _autoFilter = autoFilter;
            _column = column;
        }

        public IXLFilteredColumn AddFilter(XLCellValue value)
        {
            _autoFilter.AddFilter(_column, XLFilter.CreateRegularFilter(value));
            _autoFilter.Reapply();
            return this;
        }
    }
}
