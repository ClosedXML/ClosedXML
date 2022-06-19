namespace ClosedXML.Excel
{
    internal class XLFilterConnector : IXLFilterConnector
    {
        private readonly XLAutoFilter _autoFilter;
        private readonly int _column;

        public XLFilterConnector(XLAutoFilter autoFilter, int column)
        {
            _autoFilter = autoFilter;
            _column = column;
        }

        #region IXLFilterConnector Members

        public IXLCustomFilteredColumn And => new XLCustomFilteredColumn(_autoFilter, _column, XLConnector.And);

        public IXLCustomFilteredColumn Or => new XLCustomFilteredColumn(_autoFilter, _column, XLConnector.Or);

        #endregion
    }
}