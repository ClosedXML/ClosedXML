using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ClosedXML.Excel
{
    using System.Linq;
    using System.Text.RegularExpressions;
    using System.Drawing;

    /// <summary>
    /// 	Common methods
    /// </summary>
    public static class XLHelper
    {
        public const int MinRowNumber = 1;
        public const int MinColumnNumber = 1;
        public const int MaxRowNumber = 1048576;
        public const int MaxColumnNumber = 16384;
        public const String MaxColumnLetter = "XFD";
        public const Double Epsilon = 1e-10;

        private const Int32 TwoT26 = 26*26;
        internal static readonly NumberFormatInfo NumberFormatForParse = CultureInfo.InvariantCulture.NumberFormat;
        internal static readonly Graphics Graphic = Graphics.FromImage(new Bitmap(200, 200));
        internal static readonly Double DpiX = Graphic.DpiX;

        internal static readonly Regex A1SimpleRegex = new Regex(
            @"\A"
            + @"(?<Reference>" // Start Group to pick
            + @"(?<Sheet>" // Start Sheet Name, optional
            + @"("
            + @"\'([^\[\]\*/\\\?:\']+|\'\')\'"
            // Sheet name with special characters, surrounding apostrophes are required
            + @"|"
            + @"\'?\w+\'?" // Sheet name with letters and numbers, surrounding apostrophes are optional
            + @")"
            + @"!)?" // End Sheet Name, optional
            + @"(?<Range>" // Start range
            + @"\$?[a-zA-Z]{1,3}\$?\d{1,7}" // A1 Address 1
            + @"(?<RangeEnd>:\$?[a-zA-Z]{1,3}\$?\d{1,7})?" // A1 Address 2, optional
            + @"|"
            + @"(?<ColumnNumbers>\$?\d{1,7}:\$?\d{1,7})" // 1:1
            + @"|"
            + @"(?<ColumnLetters>\$?[a-zA-Z]{1,3}:\$?[a-zA-Z]{1,3})" // A:A
            + @")" // End Range
            + @")" // End Group to pick
            + @"\Z"
            );

        internal static readonly Regex NamedRangeReferenceRegex =
            new Regex(@"^('?(?<Sheet>[^'!]+)'?!(?<Range>.+))|((?<Table>[^\[]+)\[(?<Column>[^\]]+)\])$",
                      RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.ExplicitCapture
                );

        /// <summary>
        /// 	Gets the column number of a given column letter.
        /// </summary>
        /// <param name="columnLetter"> The column letter to translate into a column number. </param>
        public static int GetColumnNumberFromLetter(string columnLetter)
        {
            if (columnLetter[0] <= '9')
                return Int32.Parse(columnLetter, NumberFormatForParse);

            columnLetter = columnLetter.ToUpper();
            var length = columnLetter.Length;
            if (length == 1)
                return Convert.ToByte(columnLetter[0]) - 64;
            if (length == 2)
            {
                return
                    ((Convert.ToByte(columnLetter[0]) - 64)*26) +
                    (Convert.ToByte(columnLetter[1]) - 64);
            }
            if (length == 3)
            {
                return ((Convert.ToByte(columnLetter[0]) - 64)*TwoT26) +
                       ((Convert.ToByte(columnLetter[1]) - 64)*26) +
                       (Convert.ToByte(columnLetter[2]) - 64);
            }
            throw new ApplicationException("Column Length must be between 1 and 3.");
        }

        /// <summary>
        /// 	Gets the column letter of a given column number.
        /// </summary>
        /// <param name="column"> The column number to translate into a column letter. </param>
        public static string GetColumnLetterFromNumber(int column)
        {
            #region Check

            if (column <= 0)
                throw new ArgumentOutOfRangeException("column", "Must be more than 0");

            #endregion

            var value = new StringBuilder(6);
            while (column > 0)
            {
                var residue = column%26;
                column /= 26;
                if (residue == 0)
                {
                    residue = 26;
                    column--;
                }
                value.Insert(0, (char) (64 + residue));
            }
            return value.ToString();
        }

        public static bool IsValidColumn(string column)
        {
            var length = column.Length;
            if (IsNullOrWhiteSpace(column) || length > 3)
                return false;

            var theColumn = column.ToUpper();


            var isValid = theColumn[0] >= 'A' && theColumn[0] <= 'Z';
            if (length == 1)
                return isValid;

            if (length == 2)
                return isValid && theColumn[1] >= 'A' && theColumn[1] <= 'Z';

            if (theColumn[0] >= 'A' && theColumn[0] < 'X')
                return theColumn[1] >= 'A' && theColumn[1] <= 'Z'
                       && theColumn[2] >= 'A' && theColumn[2] <= 'Z';

            if (theColumn[0] != 'X') return false;

            if (theColumn[1] < 'F')
                return theColumn[2] >= 'A' && theColumn[2] <= 'Z';

            if (theColumn[1] != 'F') return false;

            return theColumn[2] >= 'A' && theColumn[2] <= 'D';
        }

        public static bool IsValidRow(string rowString)
        {
            Int32 row;
            if (Int32.TryParse(rowString, out row))
                return row > 0 && row <= MaxRowNumber;
            return false;
        }

        public static bool IsValidA1Address(string address)
        {
            if (IsNullOrWhiteSpace(address))
                return false;

            address = address.Replace("$", "");
            var rowPos = 0;
            var addressLength = address.Length;
            while (rowPos < addressLength && (address[rowPos] > '9' || address[rowPos] < '0'))
                rowPos++;

            return
                rowPos < addressLength
                && IsValidRow(address.Substring(rowPos))
                && IsValidColumn(address.Substring(0, rowPos));
        }

        public static Boolean IsValidRangeAddress(String rangeAddress)
        {
            return A1SimpleRegex.IsMatch(rangeAddress);
        }

        public static int GetColumnNumberFromAddress(string cellAddressString)
        {
            var rowPos = 0;
            while (cellAddressString[rowPos] > '9')
                rowPos++;

            return GetColumnNumberFromLetter(cellAddressString.Substring(0, rowPos));
        }

        internal static string[] SplitRange(string range)
        {
            return range.Contains('-') ? range.Replace('-', ':').Split(':') : range.Split(':');
        }

        public static Int32 GetPtFromPx(Double px)
        {
            return Convert.ToInt32(px*72.0/DpiX);
        }

        public static Double GetPxFromPt(Int32 pt)
        {
            return Convert.ToDouble(pt)*DpiX/72.0;
        }

        internal static IXLTableRows InsertRowsWithoutEvents(Func<int, bool, IXLRangeRows> insertFunc,
                                                             XLTableRange tableRange, Int32 numberOfRows,
                                                             Boolean expandTable)
        {
            var ws = tableRange.Worksheet;
            var tracking = ws.EventTrackingEnabled;
            ws.EventTrackingEnabled = false;

            var rows = new XLTableRows(ws.Style);
            var inserted = insertFunc(numberOfRows, false);
            inserted.ForEach(r => rows.Add(new XLTableRow(tableRange, r as XLRangeRow)));

            if (expandTable)
                tableRange.Table.ExpandTableRows(numberOfRows);

            ws.EventTrackingEnabled = tracking;

            return rows;
        }



        public static bool IsNullOrWhiteSpace(string value)
        {
#if NET4
            return String.IsNullOrWhiteSpace(value);
#else
            if (value != null)
            {
                var length = value.Length;
                for (int i = 0; i < length; i++)
                {
                    if (!char.IsWhiteSpace(value[i]))
                    {
                        return false;
                    }
                }
            }
            return true;
#endif

        }
    }
}