// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    /// <summary>
    ///  Relative or absolute reference to a range of columns
    /// </summary>
    internal class XLColumnRangeReference : IXLCompoundReference
    {
        #region Public Properties

        public XLColumnReference FirstColumn { get; }

        public XLColumnReference LastColumn { get; }

        #endregion Public Properties

        #region Public Constructors

        public XLColumnRangeReference(XLColumnReference firstColumn, XLColumnReference lastColumn)
        {
            FirstColumn = firstColumn ?? throw new ArgumentNullException(nameof(firstColumn));
            LastColumn = lastColumn ?? throw new ArgumentNullException(nameof(lastColumn));
        }

        #endregion Public Constructors

        #region Public Methods

        public override string ToString() => ToStringR1C1();

        public string ToStringA1(IXLAddress baseAddress)
        {
            var firstColumnNumber = FirstColumn.ColumnIsAbsolute
                ? FirstColumn.Column
                : baseAddress.ColumnNumber + FirstColumn.Column;

            if (firstColumnNumber < 1 || firstColumnNumber > XLHelper.MaxColumnNumber)
                return "#REF!";

            var lastColumnNumber = LastColumn.ColumnIsAbsolute
                ? LastColumn.Column
                : baseAddress.ColumnNumber + LastColumn.Column;

            if (lastColumnNumber < 1 || lastColumnNumber > XLHelper.MaxColumnNumber)
                return "#REF!";

            var firstColumnLetter = XLHelper.GetColumnLetterFromNumber(firstColumnNumber);
            var lastColumnLetter = XLHelper.GetColumnLetterFromNumber(lastColumnNumber);

            if (FirstColumn.ColumnIsAbsolute && LastColumn.ColumnIsAbsolute)
                return $"${firstColumnLetter}:${lastColumnLetter}";

            if (FirstColumn.ColumnIsAbsolute)
                return $"${firstColumnLetter}:{lastColumnLetter}";

            if (LastColumn.ColumnIsAbsolute)
                return $"{firstColumnLetter}:${lastColumnLetter}";

            return $"{firstColumnLetter}:{lastColumnLetter}";
        }

        public string ToStringR1C1()
        {
            return $"{FirstColumn.ToStringR1C1()}:{LastColumn.ToStringR1C1()}";
        }

        #endregion Public Methods
    }
}
