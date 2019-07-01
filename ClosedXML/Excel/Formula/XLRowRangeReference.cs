// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    /// <summary>
    ///  Relative or absolute reference to a range of rows
    /// </summary>
    internal class XLRowRangeReference : IXLCompoundReference
    {
        #region Public Properties

        public XLRowReference FirstRow { get; }

        public XLRowReference LastRow { get; }

        #endregion Public Properties

        #region Public Constructors

        public XLRowRangeReference(XLRowReference firstRow, XLRowReference lastRow)
        {
            FirstRow = firstRow ?? throw new ArgumentNullException(nameof(firstRow));
            LastRow = lastRow ?? throw new ArgumentNullException(nameof(lastRow));
        }

        #endregion Public Constructors

        #region Public Methods

        public override string ToString() => ToStringR1C1();

        public string ToStringA1(IXLAddress baseAddress)
        {
            var firstRowNumber = FirstRow.RowIsAbsolute
                ? FirstRow.Row
                : baseAddress.RowNumber + FirstRow.Row;

            if (firstRowNumber < 1 || firstRowNumber > XLHelper.MaxRowNumber)
                return "#REF!";

            var lastRowNumber = LastRow.RowIsAbsolute
                ? LastRow.Row
                : baseAddress.RowNumber + LastRow.Row;

            if (lastRowNumber < 1 || lastRowNumber > XLHelper.MaxRowNumber)
                return "#REF!";

            if (FirstRow.RowIsAbsolute && LastRow.RowIsAbsolute)
                return $"${firstRowNumber}:${lastRowNumber}";

            if (FirstRow.RowIsAbsolute)
                return $"${firstRowNumber}:{lastRowNumber}";

            if (LastRow.RowIsAbsolute)
                return $"{firstRowNumber}:${lastRowNumber}";

            return $"{firstRowNumber}:{lastRowNumber}";
        }

        public string ToStringR1C1()
        {
            return $"{FirstRow.ToStringR1C1()}:{LastRow.ToStringR1C1()}";
        }

        #endregion Public Methods
    }
}
