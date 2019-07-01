// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    /// <summary>
    /// Relative or absolute reference to a single cell
    /// </summary>
    internal class XLCellReference : IXLSimpleReference
    {
        #region Public Properties

        public int Column { get; }

        public bool ColumnIsAbsolute { get; }

        public int Row { get; }
        public bool RowIsAbsolute { get; }

        #endregion Public Properties

        #region Public Constructors

        public XLCellReference(int row, int column, bool rowIsAbsolute, bool columnIsAbsolute)
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

            if (columnIsAbsolute)
            {
                if (column < 1 || column > XLHelper.MaxColumnNumber)
                    throw new ArgumentOutOfRangeException($"Column number must be between 1 and {XLHelper.MaxColumnNumber}, passed {column}");
            }
            else
            {
                if (column < -XLHelper.MaxColumnNumber || column > XLHelper.MaxColumnNumber)
                    throw new ArgumentOutOfRangeException($"Relative column number must be between -{XLHelper.MaxColumnNumber} and {XLHelper.MaxColumnNumber}, passed {column}");
            }

            Row = row;
            Column = column;
            RowIsAbsolute = rowIsAbsolute;
            ColumnIsAbsolute = columnIsAbsolute;
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

            var rowPart = RowIsAbsolute
                ? $"${rowNumber}"
                : rowNumber.ToString();

            var columnNumber = ColumnIsAbsolute
                ? Column
                : baseAddress.ColumnNumber + Column;

            if (columnNumber < 1 || columnNumber > XLHelper.MaxColumnNumber)
                return "#REF!";

            var columnLetter = XLHelper.GetColumnLetterFromNumber(columnNumber);
            var columnPart = ColumnIsAbsolute
                ? $"${columnLetter}"
                : columnLetter;

            return columnPart + rowPart;
        }

        public string ToStringR1C1()
        {
            string rowPart;
            if (RowIsAbsolute)
                rowPart = $"R{Row}";
            else if (Row == 0)
                rowPart = "R";
            else
                rowPart = $"R[{Row}]";

            string columnPart;
            if (ColumnIsAbsolute)
                columnPart = $"C{Column}";
            else if (Column == 0)
                columnPart = "C";
            else
                columnPart = $"C[{Column}]";

            return rowPart + columnPart;
        }

        #endregion Public Methods
    }
}
