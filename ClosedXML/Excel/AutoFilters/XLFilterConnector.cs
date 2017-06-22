using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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

        #region IXLFilterConnector Members

        public IXLCustomFilteredColumn And
        {
            get { return new XLCustomFilteredColumn(_autoFilter, _column, XLConnector.And); }
        }

        public IXLCustomFilteredColumn Or
        {
            get { return new XLCustomFilteredColumn(_autoFilter, _column, XLConnector.Or); }
        }

        #endregion
    }
}