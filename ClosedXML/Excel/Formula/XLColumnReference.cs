// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    /// <summary>
    ///  Relative or absolute reference to a single column
    /// </summary>
    internal class XLColumnReference : IXLSimpleReference
    {
        #region Public Properties

        public int Column { get; }

        public bool ColumnIsAbsolute { get; }

        #endregion Public Properties

        #region Public Constructors

        public XLColumnReference(int column, bool columnIsAbsolute)
        {
            if (columnIsAbsolute)
            {
                if (column < 1 || column > XLHelper.MaxColumnNumber)
                    throw new ArgumentOutOfRangeException($"Column number must be between 1 and {XLHelper.MaxColumnNumber}, passed {column}");
            }
            else
            {
                if (column < -XLHelper.MaxColumnNumber || column > XLHelper.MaxColumnNumber)
                    throw new ArgumentOutOfRangeException($"Relative Column number must be between -{XLHelper.MaxColumnNumber} and {XLHelper.MaxColumnNumber}, passed {column}");
            }

            Column = column;
            ColumnIsAbsolute = columnIsAbsolute;
        }

        #endregion Public Constructors

        #region Public Methods

        public override string ToString() => ToStringR1C1();

        public string ToStringA1(IXLAddress baseAddress)
        {
            var columnNumber = ColumnIsAbsolute
                ? Column
                : baseAddress.ColumnNumber + Column;

            if (columnNumber < 1 || columnNumber > XLHelper.MaxColumnNumber)
                return "#REF!";

            var columnLetter = XLHelper.GetColumnLetterFromNumber(columnNumber);
            if (ColumnIsAbsolute)
                return $"${columnLetter}:${columnLetter}";

            return $"{columnLetter}:{columnLetter}";
        }

        public string ToStringR1C1()
        {
            if (ColumnIsAbsolute)
                return $"C{Column}";

            if (Column == 0)
                return "C";

            return $"C[{Column}]";
        }

        #endregion Public Methods
    }
}
