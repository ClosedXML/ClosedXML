// Keep this file CodeMaid organised and cleaned
using System;
using System.Text.RegularExpressions;

namespace ClosedXML.Excel
{
    internal class XLFormulaDefinitionA1Parser : XLFormulaDefinitionParserBase
    {
        #region Protected Properties

        protected override Regex ReferenceRegex => A1Regex;

        #endregion Protected Properties

        #region Protected Methods

        protected override IXLSimpleReference ParseSimpleReference(string simpleReferenceString, IXLAddress baseAddress)
        {
            var columnPart = "";
            var rowPart = "";
            var chars = simpleReferenceString.ToCharArray();
            var i = 0;
            var rowIsAbsolute = false;
            var columnIsAbsolute = false;
            while (i < chars.Length && (chars[i] == '$' || chars[i] > '9'))
            {
                if (chars[i] == '$')
                {
                    if (columnIsAbsolute || columnPart != "")
                        rowIsAbsolute = true;
                    else
                        columnIsAbsolute = true;
                }
                else
                {
                    columnPart += chars[i]; // In most cases there are only 1-2 characters, 3 at max.
                                            // Using StringBuilder would be more expensive
                }

                i++;
            }

            if (columnPart == "" && columnIsAbsolute)
            {
                rowIsAbsolute = true;
                columnIsAbsolute = false;
            }

            rowPart = simpleReferenceString.Substring(i);

            int? rowNumber = null;
            int? columnNumber = null;

            if (columnPart != "")
            {
                columnNumber = XLHelper.GetColumnNumberFromLetter(columnPart);
                if (!columnIsAbsolute)
                    columnNumber -= baseAddress.ColumnNumber;
            }

            if (rowPart != "")
            {
                rowNumber = int.Parse(rowPart);
                if (!rowIsAbsolute)
                    rowNumber -= baseAddress.RowNumber;
            }

            if (rowNumber.HasValue && columnNumber.HasValue)
                return new XLCellReference(rowNumber.Value, columnNumber.Value, rowIsAbsolute, columnIsAbsolute);

            if (rowNumber.HasValue)
                return new XLRowReference(rowNumber.Value, rowIsAbsolute);

            if (columnNumber.HasValue)
                return new XLColumnReference(columnNumber.Value, columnIsAbsolute);

            throw new InvalidOperationException($"Could not parse {simpleReferenceString} as a simple A1 address");
        }

        #endregion Protected Methods

        #region Private Fields

        /// <summary> A1 or A1:A1 </summary>
        private const string CellReferenceA1Template = @"(?<=(\W|^))(\$?[a-zA-Z]{1,3}\$?\d{1,7})(?::(\$?[a-zA-Z]{1,3}\$?\d{1,7}))?(?=(\W|$))";

        /// <summary> 1:1 </summary>
        private const string ColumnRangeReferenceA1Template = @"(?<=(\W|^))(\$?\d{1,7}:\$?\d{1,7})(?=(\W|$))";

        /// <summary> A:A </summary>
        private const string RowRangeReferenceA1Template = @"(?<=(\W|^))(\$?[a-zA-Z]{1,3}:\$?[a-zA-Z]{1,3})(?=(\W|$))";

        private static readonly Regex A1Regex =
            new Regex($"{CellReferenceA1Template}|{ColumnRangeReferenceA1Template}|{RowRangeReferenceA1Template}",
                RegexOptions.Compiled);

        #endregion Private Fields
    }
}
