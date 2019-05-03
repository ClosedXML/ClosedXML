using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ClosedXML.Excel
{
    internal partial class XLCell
    {

        private string GetFormula(string strValue, FormulaConversionType conversionType, int rowsToShift,
            int columnsToShift)
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
                        ? GetR1C1Address(matchString, rowsToShift, columnsToShift)
                        : GetA1Address(matchString, rowsToShift, columnsToShift));
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

        private string GetA1Address(string r1C1Address, int rowsToShift, int columnsToShift)
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
                    leftPart = GetA1Row(p1, rowsToShift);
                    rightPart = GetA1Row(p2, rowsToShift);
                }
                else
                {
                    leftPart = GetA1Column(p1, columnsToShift);
                    rightPart = GetA1Column(p2, columnsToShift);
                }

                return leftPart + ":" + rightPart;
            }

            try
            {
                var rowPart = addressToUse.Substring(0, addressToUse.IndexOf("C"));
                var rowToReturn = GetA1Row(rowPart, rowsToShift);

                var columnPart = addressToUse.Substring(addressToUse.IndexOf("C"));
                var columnToReturn = GetA1Column(columnPart, columnsToShift);

                var retAddress = columnToReturn + rowToReturn;
                return retAddress;
            }
            catch (ArgumentOutOfRangeException)
            {
                return "#REF!";
            }
        }

        private string GetA1Column(string columnPart, int columnsToShift)
        {
            string columnToReturn;
            if (columnPart == "C")
                columnToReturn = XLHelper.GetColumnLetterFromNumber(_columnNumber + columnsToShift);
            else
            {
                var bIndex = columnPart.IndexOf("[");
                var mIndex = columnPart.IndexOf("-");
                if (bIndex >= 0)
                {
                    columnToReturn = XLHelper.GetColumnLetterFromNumber(
                        _columnNumber +
                        Int32.Parse(columnPart.Substring(bIndex + 1, columnPart.Length - bIndex - 2)) + columnsToShift
                        );
                }
                else if (mIndex >= 0)
                {
                    columnToReturn = XLHelper.GetColumnLetterFromNumber(
                        _columnNumber + Int32.Parse(columnPart.Substring(mIndex)) + columnsToShift
                        );
                }
                else
                {
                    columnToReturn = "$" +
                                     XLHelper.GetColumnLetterFromNumber(Int32.Parse(columnPart.Substring(1)) +
                                                                        columnsToShift);
                }
            }

            return columnToReturn;
        }

        private string GetA1Row(string rowPart, int rowsToShift)
        {
            string rowToReturn;
            if (rowPart == "R")
                rowToReturn = (_rowNumber + rowsToShift).ToString();
            else
            {
                var bIndex = rowPart.IndexOf("[");
                if (bIndex >= 0)
                {
                    rowToReturn =
                        (_rowNumber + Int32.Parse(rowPart.Substring(bIndex + 1, rowPart.Length - bIndex - 2)) +
                         rowsToShift).ToString();
                }
                else
                    rowToReturn = "$" + (Int32.Parse(rowPart.Substring(1)) + rowsToShift);
            }

            return rowToReturn;
        }

        private string GetR1C1Row(int rowNumber, bool fixedRow, int rowsToShift)
        {
            string rowPart;
            rowNumber += rowsToShift;
            var rowDiff = rowNumber - _rowNumber;
            if (rowDiff != 0 || fixedRow)
                rowPart = fixedRow ? "R" + rowNumber : "R[" + rowDiff + "]";
            else
                rowPart = "R";

            return rowPart;
        }

        private string GetR1C1Column(int columnNumber, bool fixedColumn, int columnsToShift)
        {
            string columnPart;
            columnNumber += columnsToShift;
            var columnDiff = columnNumber - _columnNumber;
            if (columnDiff != 0 || fixedColumn)
                columnPart = fixedColumn ? "C" + columnNumber : "C[" + columnDiff + "]";
            else
                columnPart = "C";

            return columnPart;
        }


        private string GetR1C1Address(string a1Address, int rowsToShift, int columnsToShift)
        {
            if (a1Address.Contains(':'))
            {
                var parts = a1Address.Split(':');
                var p1 = parts[0];
                var p2 = parts[1];
                if (Int32.TryParse(p1.Replace("$", string.Empty), out Int32 row1))
                {
                    var row2 = Int32.Parse(p2.Replace("$", string.Empty));
                    var leftPart = GetR1C1Row(row1, p1.Contains('$'), rowsToShift);
                    var rightPart = GetR1C1Row(row2, p2.Contains('$'), rowsToShift);
                    return leftPart + ":" + rightPart;
                }
                else
                {
                    var column1 = XLHelper.GetColumnNumberFromLetter(p1.Replace("$", string.Empty));
                    var column2 = XLHelper.GetColumnNumberFromLetter(p2.Replace("$", string.Empty));
                    var leftPart = GetR1C1Column(column1, p1.Contains('$'), columnsToShift);
                    var rightPart = GetR1C1Column(column2, p2.Contains('$'), columnsToShift);
                    return leftPart + ":" + rightPart;
                }
            }

            var address = XLAddress.Create(Worksheet, a1Address);

            var rowPart = GetR1C1Row(address.RowNumber, address.FixedRow, rowsToShift);
            var columnPart = GetR1C1Column(address.ColumnNumber, address.FixedColumn, columnsToShift);

            return rowPart + columnPart;
        }

    }
}
