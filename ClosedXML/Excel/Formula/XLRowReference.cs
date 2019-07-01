// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    /// <summary>
    ///  Relative or absolute reference to a single row
    /// </summary>
    internal class XLRowReference : IXLSimpleReference
    {
        #region Public Properties

        public int Row { get; }

        public bool RowIsAbsolute { get; }

        #endregion Public Properties

        #region Public Constructors

        public XLRowReference(int row, bool rowIsAbsolute)
        {
            if (rowIsAbsolute)
            {
                if (row < 1 || row > XLHelper.MaxRowNumber)
                    throw new ArgumentOutOfRangeException($"Row number must be between 1 and {XLHelper.MaxRowNumber}, passed {row}");
            }
            else
            {
                if (row < -XLHelper.MaxRowNumber || row > XLHelper.MaxRowNumber)
                    throw new ArgumentOutOfRangeException($"Relative row number must be between -{XLHelper.MaxRowNumber} and {XLHelper.MaxRowNumber}, passed {row}");
            }

            Row = row;
            RowIsAbsolute = rowIsAbsolute;
        }

        #endregion Public Constructors

        #region Public Methods

        public override string ToString() => ToStringR1C1();

        public string ToStringA1(IXLAddress baseAddress)
        {
            var rowNumber = RowIsAbsolute
                ? Row
                : baseAddress.RowNumber + Row;

            if (rowNumber < 1 || rowNumber > XLHelper.MaxRowNumber)
                return "#REF!";

            if (RowIsAbsolute)
                return $"${rowNumber}:${rowNumber}";

            return $"{rowNumber}:{rowNumber}";
        }

        public string ToStringR1C1()
        {
            if (RowIsAbsolute)
                return $"R{Row}";

            if (Row == 0)
                return "R";

            return $"R[{Row}]";
        }

        #endregion Public Methods
    }
}
