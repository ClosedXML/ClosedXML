using System;

namespace ClosedXML.Excel
{
    internal class XLRangeFactory
    {
        private readonly XLWorksheet _worksheet;

        public XLRangeFactory(XLWorksheet worksheet)
        {
            _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        }

        #region Methods

        public XLRangeBase Create(XLRangeKey key)
        {
            switch (key.RangeType)
            {
                case XLRangeType.Range:
                    return CreateRange(key.RangeAddress);

                case XLRangeType.Column:
                    return CreateColumn(key.RangeAddress.FirstAddress.ColumnNumber);

                case XLRangeType.Row:
                    return CreateColumn(key.RangeAddress.FirstAddress.RowNumber);

                case XLRangeType.RangeColumn:
                    return CreateRangeColumn(key.RangeAddress);

                case XLRangeType.RangeRow:
                    return CreateRangeRow(key.RangeAddress);

                case XLRangeType.Table:
                    return CreateTable(key.RangeAddress);

                case XLRangeType.Worksheet:
                default:
                    throw new NotImplementedException(key.RangeType.ToString());
            }
        }

        public XLRange CreateRange(XLRangeAddress rangeAddress)
        {
            return new XLRange(rangeAddress, _worksheet.Style);
        }

        public XLColumn CreateColumn(int columnNumber)
        {
            return new XLColumn(_worksheet, columnNumber);
        }

        public XLRow CreateRow(int rowNumber)
        {
            return new XLRow(_worksheet, rowNumber);
        }

        public XLRangeColumn CreateRangeColumn(XLRangeAddress rangeAddress)
        {
            return new XLRangeColumn(rangeAddress, _worksheet.Style);
        }

        public XLRangeRow CreateRangeRow(XLRangeAddress rangeAddress)
        {
            return new XLRangeRow(rangeAddress, _worksheet.Style);
        }

        public XLTable CreateTable(XLRangeAddress rangeAddress)
        {
            return new XLTable(rangeAddress, _worksheet.Style);
        }

        #endregion Methods
    }
}
