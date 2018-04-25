namespace ClosedXML.Excel
{
    internal class XLRangeParameters
    {
        #region Constructor

        public XLRangeParameters(XLRangeAddress rangeAddress, IXLStyle defaultStyle)
        {
            RangeAddress = rangeAddress;
            DefaultStyle = defaultStyle;
        }

        #endregion

        #region Properties

        public XLRangeAddress RangeAddress { get; private set; }

        public IXLStyle DefaultStyle { get; private set; }
        #endregion
    }
}
