using System;
using System.Globalization;

namespace ClosedXML.Excel
{
    using System.Drawing;
    using System.Linq;
    using System.Text.RegularExpressions;

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

        private const Int32 TwoT26 = 26 * 26;
        internal static readonly Graphics Graphic = Graphics.FromImage(new Bitmap(200, 200));
        internal static readonly Double DpiX = Graphic.DpiX;
        internal static readonly NumberStyles NumberStyle = NumberStyles.AllowDecimalPoint | NumberStyles.AllowLeadingSign | NumberStyles.AllowLeadingWhite | NumberStyles.AllowTrailingWhite | NumberStyles.AllowExponent;
        internal static readonly CultureInfo ParseCulture = CultureInfo.InvariantCulture;

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
            , RegexOptions.Compiled);

        internal static readonly Regex NamedRangeReferenceRegex =
            new Regex(@"^('?(?<Sheet>[^'!]+)'?!(?<Range>.+))|((?<Table>[^\[]+)\[(?<Column>[^\]]+)\])$",
                      RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.ExplicitCapture
                );

        /// <summary>
        /// Gets the column number of a given column letter.
        /// </summary>
        /// <param name="columnLetter"> The column letter to translate into a column number. </param>
        public static int GetColumnNumberFromLetter(string columnLetter)
        {
            if (string.IsNullOrEmpty(columnLetter)) throw new ArgumentNullException("columnLetter");

            int retVal;
            columnLetter = columnLetter.ToUpper();

            //Extra check because we allow users to pass row col positions in as strings
            if (columnLetter[0] <= '9')
            {
                retVal = Int32.Parse(columnLetter, XLHelper.NumberStyle, XLHelper.ParseCulture);
                return retVal;
            }

            int sum = 0;

            for (int i = 0; i < columnLetter.Length; i++)
            {
                sum *= 26;
                sum += (columnLetter[i] - 'A' + 1);
            }

            return sum;
        }

        private static readonly string[] letters = new[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

        /// <summary>
        /// Gets the column letter of a given column number.
        /// </summary>
        /// <param name="columnNumber">The column number to translate into a column letter.</param>
        /// <param name="trimToAllowed">if set to <c>true</c> the column letter will be restricted to the allowed range.</param>
        /// <returns></returns>
        public static string GetColumnLetterFromNumber(int columnNumber, bool trimToAllowed = false)
        {
            if (trimToAllowed) columnNumber = TrimColumnNumber(columnNumber);

            columnNumber--; // Adjust for start on column 1
            if (columnNumber <= 25)
            {
                return letters[columnNumber];
            }
            var firstPart = (columnNumber) / 26;
            var remainder = ((columnNumber) % 26) + 1;
            return GetColumnLetterFromNumber(firstPart) + GetColumnLetterFromNumber(remainder);
        }

        internal static int TrimColumnNumber(int columnNumber)
        {
            return Math.Max(XLHelper.MinColumnNumber, Math.Min(XLHelper.MaxColumnNumber, columnNumber));
        }

        internal static int TrimRowNumber(int rowNumber)
        {
            return Math.Max(XLHelper.MinRowNumber, Math.Min(XLHelper.MaxRowNumber, rowNumber));
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

        public static Boolean IsValidRangeAddress(IXLRangeAddress rangeAddress)
        {
            return !rangeAddress.IsInvalid
                   && rangeAddress.FirstAddress.RowNumber >= 1 && rangeAddress.LastAddress.RowNumber <= MaxRowNumber
                   && rangeAddress.FirstAddress.ColumnNumber >= 1 && rangeAddress.LastAddress.ColumnNumber <= MaxColumnNumber
                   && rangeAddress.FirstAddress.RowNumber <= rangeAddress.LastAddress.RowNumber
                   && rangeAddress.FirstAddress.ColumnNumber <= rangeAddress.LastAddress.ColumnNumber;
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
            return Convert.ToInt32(px * 72.0 / DpiX);
        }

        public static Double GetPxFromPt(Int32 pt)
        {
            return Convert.ToDouble(pt) * DpiX / 72.0;
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

        private static readonly Regex A1RegexRelative = new Regex(
      @"(?<=\W)(?<one>\$?[a-zA-Z]{1,3}\$?\d{1,7})(?=\W)" // A1
    + @"|(?<=\W)(?<two>\$?\d{1,7}:\$?\d{1,7})(?=\W)" // 1:1
    + @"|(?<=\W)(?<three>\$?[a-zA-Z]{1,3}:\$?[a-zA-Z]{1,3})(?=\W)", RegexOptions.Compiled); // A:A

        private static string Evaluator(Match match, Int32 row, String column)
        {
            if (match.Groups["one"].Success)
            {
                var split = match.Groups["one"].Value.Split('$');
                if (split.Length == 1) return column + row; // A1
                if (split.Length == 3) return match.Groups["one"].Value; // $A$1
                var a = XLAddress.Create(match.Groups["one"].Value);
                if (split[0] == String.Empty) return "$" + a.ColumnLetter + row; // $A1
                return column + "$" + a.RowNumber;
            }

            if (match.Groups["two"].Success)
                return ReplaceGroup(match.Groups["two"].Value, row.ToString());

            return ReplaceGroup(match.Groups["three"].Value, column);
        }

        private static String ReplaceGroup(String value, String item)
        {
            var split = value.Split(':');
            String ret1 = split[0].StartsWith("$") ? split[0] : item;
            String ret2 = split[1].StartsWith("$") ? split[1] : item;
            return ret1 + ":" + ret2;
        }

        internal static String ReplaceRelative(String value, Int32 row, String column)
        {
            var oldValue = ">" + value + "<";
            var newVal = A1RegexRelative.Replace(oldValue, m => Evaluator(m, row, column));
            return newVal.Substring(1, newVal.Length - 2);
        }

        public static Boolean AreEqual(Double d1, Double d2)
        {
            return Math.Abs(d1 - d2) < Epsilon;
        }

        public static DateTime GetDate(Object v)
        {
            // handle dates
            if (v is DateTime)
            {
                return (DateTime)v;
            }

            // handle doubles
            if (v is double && ((double)v).IsValidOADateNumber())
            {
                return DateTime.FromOADate((double)v);
            }

            // handle everything else
            return (DateTime)Convert.ChangeType(v, typeof(DateTime));
        }

        internal static bool IsValidOADateNumber(this double d)
        {
            return -657435 <= d && d < 2958466;
        }
    }
}
