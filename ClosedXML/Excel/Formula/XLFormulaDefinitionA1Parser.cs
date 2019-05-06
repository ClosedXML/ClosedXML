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

        protected override IXLSimpleReference ParseSimpleReference(string simpleReferenceString)
        {
            throw new NotImplementedException();
        }

        #endregion Protected Methods

        #region Private Fields

        /// <summary> A1 </summary>
        private const string CellReferenceA1Template = @"(?<=(\W|^))(\$?[a-zA-Z]{1,3}\$?\d{1,7})(?=(\W|$))";

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
