using System;

namespace ClosedXML.Excel
{
    internal class XLRangeFactory
    {
        #region Properties

        public XLWorksheet Worksheet { get; private set; }

        #endregion Properties

        #region Constructors

        public XLRangeFactory(XLWorksheet worksheet)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));
            Worksheet = worksheet;
        }

        #endregion Constructors

        #region Methods

        public XLRangeBase Create(XLRangeKey key)
        {
            switch (key.RangeType)
            {
                case XLRangeType.Range:
                    return CreateRange(key.RangeAddress);
                    break;

                case XLRangeType.Column:
                    return CreateColumn(key.RangeAddress.FirstAddress.ColumnNumber);
                    break;

                case XLRangeType.Row:
                    return CreateColumn(key.RangeAddress.FirstAddress.RowNumber);
                    break;

                case XLRangeType.RangeColumn:
                //break;
                case XLRangeType.RangeRow:
                //break;
                case XLRangeType.Table:
                //break;
                case XLRangeType.Worksheet:
                //break;
                default:
                    throw new NotImplementedException(key.RangeType.ToString());
                    break;
            }
        }

        public XLColumn CreateColumn(int columnNumber)
        {
            return new XLColumn(Worksheet, columnNumber);
        }

        public XLRange CreateRange(XLRangeAddress rangeAddress)
        {
            var xlRangeParameters = new XLRangeParameters(rangeAddress, Worksheet.Style);
            return new XLRange(xlRangeParameters);
        }

        public XLRow CreateRow(int rowNumber)
        {
            return new XLRow(Worksheet, rowNumber);
        }

        #endregion Methods
    }
}
