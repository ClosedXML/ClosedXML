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
            var xlRangeParameters = new XLRangeParameters(rangeAddress, Worksheet.Style);
            return new XLRange(xlRangeParameters);
        }

        public XLColumn CreateColumn(int columnNumber)
        {
            return new XLColumn(Worksheet, columnNumber);
        }

        public XLRow CreateRow(int rowNumber)
        {
            return new XLRow(Worksheet, rowNumber);
        }

        public XLRangeColumn CreateRangeColumn(XLRangeAddress rangeAddress)
        {
            var xlRangeParameters = new XLRangeParameters(rangeAddress, Worksheet.Style);
            return new XLRangeColumn(xlRangeParameters);
        }

        public XLRangeRow CreateRangeRow(XLRangeAddress rangeAddress)
        {
            var xlRangeParameters = new XLRangeParameters(rangeAddress, Worksheet.Style);
            return new XLRangeRow(xlRangeParameters);
        }

        public XLTable CreateTable(XLRangeAddress rangeAddress)
        {
            return new XLTable(new XLRangeParameters(rangeAddress, Worksheet.Style));
        }

        #endregion Methods
    }
}
