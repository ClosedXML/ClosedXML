using System;

namespace ClosedXML.Excel
{
    internal class XLFilterConnector : IXLFilterConnector
    {
        private readonly XLAutoFilter _autoFilter;
        private readonly Int32 _column;

        public XLFilterConnector(XLAutoFilter autoFilter, Int32 column)
        {
            _autoFilter = autoFilter;
            _column = column;
        }

        public IXLCustomFilteredColumn And => new XLCustomFilteredColumn(_autoFilter, _column, XLConnector.And);

        public IXLCustomFilteredColumn Or => new XLCustomFilteredColumn(_autoFilter, _column, XLConnector.Or);
    }
}
