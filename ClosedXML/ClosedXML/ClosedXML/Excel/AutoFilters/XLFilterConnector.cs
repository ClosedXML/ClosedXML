using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLFilterConnector: IXLFilterConnector
    {
        XLAutoFilter _autoFilter;
        Int32 _column;
        public XLFilterConnector(XLAutoFilter autoFilter, Int32 column)
        {
            _autoFilter = autoFilter;
            _column = column;
        }
        public IXLCustomFilteredColumn And { get { return new XLCustomFilteredColumn(_autoFilter, _column, XLConnector.And); } }
        public IXLCustomFilteredColumn Or { get { return new XLCustomFilteredColumn(_autoFilter, _column, XLConnector.Or); } }
    }
}
