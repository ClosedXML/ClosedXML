// Keep this file CodeMaid organised and cleaned
using System;
using System.Text.RegularExpressions;

namespace ClosedXML.Excel
{
    internal class XLFormulaDefinitionR1C1Parser : XLFormulaDefinitionParserBase
    {
        #region Protected Properties

        protected override Regex ReferenceRegex => R1C1Regex;

        #endregion Protected Properties

        #region Protected Methods

        protected override IXLSimpleReference ParseSimpleReference(string simpleReferenceString)
        {
            var rowPartLength = simpleReferenceString.IndexOf("C", StringComparison.OrdinalIgnoreCase);

            if (rowPartLength < 0)
                rowPartLength = simpleReferenceString.Length;

            int? rowNumber = null;
            int? columnNumber = null;
            var rowIsAbsolute = false;
            var columnIsAbsolute = false;

            if (rowPartLength > 0)
            {
                var rowPart = simpleReferenceString.Substring(0, rowPartLength);
                rowIsAbsolute = !rowPart.Contains("[");
                rowPart = rowPart.Trim(R1C1TrimChars);
                if (rowPart == string.Empty)
                {
                    rowNumber = 0;
                    rowIsAbsolute = false;
                }
                else
                {
                    rowNumber = int.Parse(rowPart);
                }
            }

            if (rowPartLength < simpleReferenceString.Length)
            {
                var columnPart = simpleReferenceString.Substring(rowPartLength);
                columnIsAbsolute = !columnPart.Contains("[");
                columnPart = columnPart.Trim(R1C1TrimChars);
                if (columnPart == string.Empty)
                {
                    columnNumber = 0;
                    columnIsAbsolute = false;
                }
                else
                {
                    columnNumber = int.Parse(columnPart);
                }
            }

            if (rowNumber.HasValue && columnNumber.HasValue)
                return new XLCellReference(rowNumber.Value, columnNumber.Value, rowIsAbsolute, columnIsAbsolute);

            if (rowNumber.HasValue)
                return new XLRowReference(rowNumber.Value, rowIsAbsolute);

            if (columnNumber.HasValue)
                return new XLColumnReference(columnNumber.Value, columnIsAbsolute);

            throw new InvalidOperationException($"Could not parse {simpleReferenceString} as a simple R1C1 address");
        }

        #endregion Protected Methods

        #region Private Fields

        /// <summary> C or C:C </summary>
        private const string ColumnRangeReferenceR1C1Template =
            @"(?<=([^\w:]|^))([Cc]\[?-?\d{0,5}\]?)(?::([Cc]\[?-?\d{0,5}\]?))?(?=([^\w:\-\[\]]|$))";

        /// <summary> R1C1 or R1C1:R1C1 </summary>
        private const string RangeReferenceR1C1Template =
            @"(?<=(\W|^))([Rr](?:\[-?\d{0,7}\]|\d{0,7})?[Cc](?:\[-?\d{0,5}\]|\d{0,5})?)(?::([Rr](?:\[-?\d{0,7}\]|\d{0,7})?[Cc](?:\[-?\d{0,5}\]|\d{0,5})?))?(?=(\W|$))";

        /// <summary> R or R:R </summary>
        private const string RowRangeReferenceR1C1Template =
            @"(?<=([^\w:]|^))([Rr]\[?-?\d{0,7}\]?)(?::([Rr]\[?-?\d{0,7}\]?))?(?=([^\w:\-\[\]]|$))";

        private static readonly Regex R1C1Regex =
            new Regex($"{RangeReferenceR1C1Template}" +
                $"|{ColumnRangeReferenceR1C1Template}" +
                $"|{RowRangeReferenceR1C1Template}",
                RegexOptions.Compiled);

        private static readonly char[] R1C1TrimChars = "RC[]".ToCharArray();

        #endregion Private Fields
    }
}
