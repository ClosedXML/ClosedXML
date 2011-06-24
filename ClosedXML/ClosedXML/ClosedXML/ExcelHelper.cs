using System;
using System.Globalization;
using System.Text;
using ClosedXML.Excel;

namespace ClosedXML
{
    /// <summary>
    /// Common methods
    /// </summary>
    public static class ExcelHelper
    {
        internal static readonly NumberFormatInfo NumberFormatForParse = CultureInfo.InvariantCulture.NumberFormat;
        private const Int32 TwoT26 = 26 * 26;
        /// <summary>
        /// 	Gets the column number of a given column letter.
        /// </summary>
        /// <param name = "columnLetter">The column letter to translate into a column number.</param>
        public static int GetColumnNumberFromLetter(string columnLetter)
        {
            if (columnLetter[0] <= '9')
            {
                return Int32.Parse(columnLetter, NumberFormatForParse);
            }

            columnLetter = columnLetter.ToUpper();
            var length = columnLetter.Length;
            if (length == 1)
            {
                return Convert.ToByte(columnLetter[0]) - 64;
            }
            if (length == 2)
            {
                return
                        ((Convert.ToByte(columnLetter[0]) - 64) * 26) +
                        (Convert.ToByte(columnLetter[1]) - 64);
            }
            if (length == 3)
            {
                return ((Convert.ToByte(columnLetter[0]) - 64) * TwoT26) +
                       ((Convert.ToByte(columnLetter[1]) - 64) * 26) +
                       (Convert.ToByte(columnLetter[2]) - 64);
            }
            throw new ApplicationException("Column Length must be between 1 and 3.");
        }
        /// <summary>
        /// 	Gets the column letter of a given column number.
        /// </summary>
        /// <param name = "column">The column number to translate into a column letter.</param>
        public static string GetColumnLetterFromNumber(int column)
        {
            #region Check
            if (column <= 0)
            {
                throw new ArgumentOutOfRangeException("column", "Must be more than 0");
            }
            #endregion
            var value = new StringBuilder(6);
            while (column > 0)
            {
                int residue = column % 26;
                column /= 26;
                if (residue == 0)
                {
                    residue = 26;
                    column--;
                }
                value.Insert(0, (char)(64 + residue));
            }
            return value.ToString();
        }

        public static bool IsValidColumn(string column)
        {
            if (StringExtensions.IsNullOrWhiteSpace(column) || column.Length > 3)
            {
                return false;
            }

            Boolean retVal = true;
            String theColumn = column.ToUpper();
            for (Int32 i = 0; i < column.Length; i++)
            {
                if (theColumn[i] < 'A' || theColumn[i] > 'Z' || (i == 2 && theColumn[i] > 'D'))
                {
                    retVal = false;
                    break;
                }
            }
            return retVal;
        }

        public static bool IsValidRow(string rowString)
        {
            Int32 row;
            if (Int32.TryParse(rowString, out row))
            {
                return row > 0 && row <= XLWorksheet.MaxNumberOfRows;
            }
            return false;
        }

        public static bool IsValidA1Address(string address)
        {
            address = address.Replace("$", "");
            Int32 rowPos = 0;
            Int32 addressLength = address.Length;
            while (rowPos < addressLength && (address[rowPos] > '9' || address[rowPos] < '0'))
            {
                rowPos++;
            }

            return
                    rowPos < addressLength
                    && IsValidRow(address.Substring(rowPos))
                    && IsValidColumn(address.Substring(0, rowPos));
        }

        public static int GetRowFromAddress1(string cellAddressString)
        {
            Int32 rowPos = 1;
            while (cellAddressString[rowPos] > '9')
            {
                rowPos++;
            }

            return Int32.Parse(cellAddressString.Substring(rowPos), NumberFormatForParse);
        }

        public static int GetColumnNumberFromAddress1(string cellAddressString)
        {
            Int32 rowPos = 0;
            while (cellAddressString[rowPos] > '9')
            {
                rowPos++;
            }

            return GetColumnNumberFromLetter(cellAddressString.Substring(0, rowPos));
        }

        public static int GetRowFromAddress2(string cellAddressString)
        {
            Int32 rowPos = 1;
            while (cellAddressString[rowPos] > '9')
            {
                rowPos++;
            }

            if (cellAddressString[rowPos] == '$')
            {
                return Int32.Parse(cellAddressString.Substring(rowPos + 1), NumberFormatForParse);
            }
            return Int32.Parse(cellAddressString.Substring(rowPos), NumberFormatForParse);
        }

        public static int GetColumnNumberFromAddress2(string cellAddressString)
        {
            Int32 startPos;
            if (cellAddressString[0] == '$')
            {
                startPos = 1;
            }
            else
            {
                startPos = 0;
            }

            Int32 rowPos = startPos;
            while (cellAddressString[rowPos] > '9')
            {
                rowPos++;
            }

            if (cellAddressString[rowPos] == '$')
            {
                return GetColumnNumberFromLetter(cellAddressString.Substring(startPos, rowPos - 1));
            }
            return GetColumnNumberFromLetter(cellAddressString.Substring(startPos, rowPos));
        }
    }
}