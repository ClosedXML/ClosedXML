using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ClosedXML.Excel
{
    internal enum FormulaType : byte
    {
        Normal,
        Array,
        DataTable,
        Shared // Not used
    }

    internal enum FormulaConversionType
    {
        A1ToR1C1,
        R1C1ToA1
    };

    /// <summary>
    /// A representation of a cell formula, not the formula itself (i.e. the tree).
    /// </summary>
    internal class XLCellFormula
    {
        private static readonly Regex A1Regex = new(
            @"(?<=\W)(\$?[a-zA-Z]{1,3}\$?\d{1,7})(?=\W)" // A1
            + @"|(?<=\W)(\$?\d{1,7}:\$?\d{1,7})(?=\W)" // 1:1
            + @"|(?<=\W)(\$?[a-zA-Z]{1,3}:\$?[a-zA-Z]{1,3})(?=\W)", RegexOptions.Compiled); // A:A

        private static readonly Regex R1C1Regex = new(
            @"(?<=\W)([Rr](?:\[-?\d{0,7}\]|\d{0,7})?[Cc](?:\[-?\d{0,7}\]|\d{0,7})?)(?=\W)" // R1C1
            + @"|(?<=\W)([Rr]\[?-?\d{0,7}\]?:[Rr]\[?-?\d{0,7}\]?)(?=\W)" // R:R
            + @"|(?<=\W)([Cc]\[?-?\d{0,5}\]?:[Cc]\[?-?\d{0,5}\]?)(?=\W)", RegexOptions.Compiled); // C:C

        private XLSheetPoint _input1;
        private XLSheetPoint _input2;
        private FormulaFlags _flags;
        private FormulaType _type;

        internal string A1 { get; set; }

        internal string R1C1 { get; set; }

        internal bool HasAnyFormula =>
            !String.IsNullOrWhiteSpace(A1) ||
            !String.IsNullOrEmpty(R1C1);

        internal FormulaType Type => _type;

        /// <summary>
        /// Get stored formula or or <c>string.Empty</c> if both A1/R1C1 are empty.
        /// Formula doesn't contain artificial equal sign.
        /// </summary>
        public string GetFormulaA1(XLSheetPoint cellAddress)
        {
            if (String.IsNullOrWhiteSpace(A1))
            {
                if (String.IsNullOrWhiteSpace(R1C1))
                {
                    return String.Empty;
                }

                A1 = GetFormula(R1C1, FormulaConversionType.R1C1ToA1, cellAddress);
            }

            if (A1.Trim()[0] == '=')
                return A1.Substring(1);

            if (A1.Trim().StartsWith("{="))
                return "{" + A1.Substring(2);

            return A1;
        }

        internal static string GetFormula(string strValue, FormulaConversionType conversionType, XLSheetPoint cellAddress)
        {
            if (String.IsNullOrWhiteSpace(strValue))
                return String.Empty;

            var value = ">" + strValue + "<";

            var regex = conversionType == FormulaConversionType.A1ToR1C1 ? A1Regex : R1C1Regex;

            var sb = new StringBuilder();
            var lastIndex = 0;

            foreach (var match in regex.Matches(value).Cast<Match>())
            {
                var matchString = match.Value;
                var matchIndex = match.Index;
                if (value.Substring(0, matchIndex).CharCount('"') % 2 == 0
                    && value.Substring(0, matchIndex).CharCount('\'') % 2 == 0)
                {
                    // Check if the match is in between quotes
                    sb.Append(value.Substring(lastIndex, matchIndex - lastIndex));
                    sb.Append(conversionType == FormulaConversionType.A1ToR1C1
                        ? GetR1C1Address(matchString, cellAddress)
                        : GetA1Address(matchString, cellAddress));
                }
                else
                    sb.Append(value.Substring(lastIndex, matchIndex - lastIndex + matchString.Length));
                lastIndex = matchIndex + matchString.Length;
            }

            if (lastIndex < value.Length)
                sb.Append(value.Substring(lastIndex));

            var retVal = sb.ToString();
            return retVal.Substring(1, retVal.Length - 2);
        }

        private static string GetA1Address(string r1C1Address, XLSheetPoint cellAddress)
        {
            var addressToUse = r1C1Address.ToUpper();

            if (addressToUse.Contains(':'))
            {
                var parts = addressToUse.Split(':');
                var p1 = parts[0];
                var p2 = parts[1];
                string leftPart;
                string rightPart;
                if (p1.StartsWith("R"))
                {
                    leftPart = GetA1Row(p1, cellAddress.Row);
                    rightPart = GetA1Row(p2, cellAddress.Row);
                }
                else
                {
                    leftPart = GetA1Column(p1, cellAddress.Column);
                    rightPart = GetA1Column(p2, cellAddress.Column);
                }

                return leftPart + ":" + rightPart;
            }

            try
            {
                var rowPart = addressToUse.Substring(0, addressToUse.IndexOf("C"));
                var rowToReturn = GetA1Row(rowPart, cellAddress.Row);

                var columnPart = addressToUse.Substring(addressToUse.IndexOf("C"));
                var columnToReturn = GetA1Column(columnPart, cellAddress.Column);

                var retAddress = columnToReturn + rowToReturn;
                return retAddress;
            }
            catch (ArgumentOutOfRangeException)
            {
                return "#REF!";
            }
        }

        private static string GetA1Column(string columnPartRC, int cellColumn)
        {
            string columnToReturn;
            if (columnPartRC == "C")
                columnToReturn = XLHelper.GetColumnLetterFromNumber(cellColumn);
            else
            {
                var bIndex = columnPartRC.IndexOf("[");
                var mIndex = columnPartRC.IndexOf("-");
                if (bIndex >= 0)
                {
                    columnToReturn = XLHelper.GetColumnLetterFromNumber(
                        cellColumn +
                        Int32.Parse(columnPartRC.Substring(bIndex + 1, columnPartRC.Length - bIndex - 2))
                    );
                }
                else if (mIndex >= 0)
                {
                    columnToReturn = XLHelper.GetColumnLetterFromNumber(
                        cellColumn + Int32.Parse(columnPartRC.Substring(mIndex))
                    );
                }
                else
                {
                    columnToReturn = "$" +
                                     XLHelper.GetColumnLetterFromNumber(Int32.Parse(columnPartRC.Substring(1)));
                }
            }

            return columnToReturn;
        }

        private static string GetA1Row(string rowPartRC, int cellRow)
        {
            string rowToReturn;
            if (rowPartRC == "R")
                rowToReturn = cellRow.ToString();
            else
            {
                var bIndex = rowPartRC.IndexOf("[");
                if (bIndex >= 0)
                {
                    rowToReturn =
                        (cellRow + Int32.Parse(rowPartRC.Substring(bIndex + 1, rowPartRC.Length - bIndex - 2))).ToString();
                }
                else
                    rowToReturn = "$" + (Int32.Parse(rowPartRC.Substring(1)));
            }

            return rowToReturn;
        }

        private static string GetR1C1Address(string a1Address, XLSheetPoint cellAddress)
        {
            if (a1Address.Contains(':'))
            {
                var parts = a1Address.Split(':');
                var p1 = parts[0];
                var p2 = parts[1];
                if (Int32.TryParse(p1.Replace("$", string.Empty), out Int32 row1))
                {
                    var row2 = Int32.Parse(p2.Replace("$", string.Empty));
                    var leftPart = GetR1C1Row(row1, p1.Contains('$'), cellAddress.Row);
                    var rightPart = GetR1C1Row(row2, p2.Contains('$'), cellAddress.Row);
                    return leftPart + ":" + rightPart;
                }
                else
                {
                    var column1 = XLHelper.GetColumnNumberFromLetter(p1.Replace("$", string.Empty));
                    var column2 = XLHelper.GetColumnNumberFromLetter(p2.Replace("$", string.Empty));
                    var leftPart = GetR1C1Column(column1, p1.Contains('$'), cellAddress.Column);
                    var rightPart = GetR1C1Column(column2, p2.Contains('$'), cellAddress.Column);
                    return leftPart + ":" + rightPart;
                }
            }

            var address = XLAddress.Create(a1Address);

            var rowPart = GetR1C1Row(address.RowNumber, address.FixedRow, cellAddress.Row);
            var columnPart = GetR1C1Column(address.ColumnNumber, address.FixedColumn, cellAddress.Column);

            return rowPart + columnPart;
        }

        private static string GetR1C1Row(int rowNumber, bool fixedRow, int cellRow)
        {
            string rowPart;
            var rowDiff = rowNumber - cellRow;
            if (rowDiff != 0 || fixedRow)
                rowPart = fixedRow ? "R" + rowNumber : "R[" + rowDiff + "]";
            else
                rowPart = "R";

            return rowPart;
        }

        private static string GetR1C1Column(int columnNumber, bool fixedColumn, int cellColumn)
        {
            string columnPart;
            var columnDiff = columnNumber - cellColumn;
            if (columnDiff != 0 || fixedColumn)
                columnPart = fixedColumn ? "C" + columnNumber : "C[" + columnDiff + "]";
            else
                columnPart = "C";

            return columnPart;
        }

        /// <summary>
        /// Set cell formula to a normal formula. Doesn't affect recalculation version.
        /// </summary>
        /// <param name="formulaA1">Doesn't start with <c>=</c>.</param>
        internal void SetNormal(string formulaA1)
        {
            A1 = formulaA1;
            R1C1 = null;
            _type = FormulaType.Normal;
            _flags = FormulaFlags.None;
        }

        /// <summary>
        /// Set cell formula to an array formula. Doesn't affect recalculation version.
        /// </summary>
        /// <param name="arrayFormulaA1">Isn't wrapped in <c>{}</c> and doesn't start with <c>=</c>.</param>
        /// <param name="aca">A flag for always calculate array. </param>
        internal void SetArray(string arrayFormulaA1, bool aca)
        {
            A1 = "{" + arrayFormulaA1 + "}";
            R1C1 = null;
            _type = FormulaType.Array;
            _flags = aca ? FormulaFlags.AlwaysCalculateArray : FormulaFlags.None;
        }

        /// <summary>
        /// Set cell formula to a 1D data table formula. Doesn't affect recalculation version.
        /// </summary>
        /// <param name="dataTableFormulaA1">Doesn't start with <c>=</c>.</param>
        /// <param name="isRowDataTable">Is data table in row (<c>true</c>) or columns (<c>false</c>)?</param>
        /// <param name="input1Address">Address of the input cell that will be replaced in the data table. If input deleted, ignored and value can be anything.</param>
        /// <param name="input1Deleted">Was the original address deleted?</param>
        internal void SetDataTable1D(
            string dataTableFormulaA1,
            bool isRowDataTable,
            XLSheetPoint input1Address,
            bool input1Deleted)
        {
            A1 = dataTableFormulaA1;
            R1C1 = null;
            _type = FormulaType.DataTable;
            _flags =
                (isRowDataTable ? FormulaFlags.Is1DRow : FormulaFlags.None) |
                (input1Deleted ? FormulaFlags.Input1Deleted : FormulaFlags.None);
            _input1 = input1Address;
        }

        /// <summary>
        /// Set cell formula to a 2D data table formula. Doesn't affect recalculation version.
        /// </summary>
        /// <param name="dataTableFormulaA1">Doesn't start with <c>=</c>.</param>
        /// <param name="input1Address">Address of the input cell that will be replaced in the data table. If input deleted, ignored and value can be anything.</param>
        /// <param name="input1Deleted">Was the original address deleted?</param>
        /// <param name="input2Address">Address of the input cell that will be replaced in the data table. If input deleted, ignored and value can be anything.</param>
        /// <param name="input2Deleted">Was the original address deleted?</param>
        internal void SetDataTable2D(
            string dataTableFormulaA1,
            XLSheetPoint input1Address,
            bool input1Deleted,
            XLSheetPoint input2Address,
            bool input2Deleted)
        {
            A1 = dataTableFormulaA1;
            R1C1 = null;
            _type = FormulaType.DataTable;
            _flags = FormulaFlags.Is2D |
                (input1Deleted ? FormulaFlags.Input1Deleted : FormulaFlags.None) |
                (input2Deleted ? FormulaFlags.Input2Deleted : FormulaFlags.None);
            _input1 = input1Address;
            _input2 = input2Address;
        }

        /// <summary>
        /// An enum to efficiently store various flags for formulas (bool takes up 1-4 bytes due to alignment).
        /// Note that each type of formula uses different flags.
        /// </summary>
        [Flags]
        private enum FormulaFlags : byte
        {
            None = 0,

            /// <summary>
            /// For Array formula. Not fully clear from documentation, but seems to be some kind of dirty flag.
            /// Current excel just writes <c>ca="1"</c> to each cell of array formula for cases described in the DOC.
            /// </summary>
            AlwaysCalculateArray = 1,

            /// <summary>
            /// For data table formula. Flag whether the data table is 2D and has two inputs.
            /// </summary>
            Is2D = 2,

            /// <summary>
            /// For data table formula. If the set, the data table is in row, not column. It uses input1 in both case, but the position
            /// is interpreted differently.
            /// </summary>
            Is1DRow = 4,

            /// <summary>
            /// For data table formula. When the input 1 cell has been deleted (not content, but the row or a column where cell was),
            /// this flag is set.
            /// </summary>
            Input1Deleted = 8,

            /// <summary>
            /// For data table formula. When the input 2 cell has been deleted (not content, but the row or a column where cell was),
            /// this flag is set.
            /// </summary>
            Input2Deleted = 16,
        }
    }
}
