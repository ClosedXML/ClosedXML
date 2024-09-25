#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ClosedXML.Excel
{
    internal static class StringExtensions
    {
        private static readonly Regex RegexNewLine = new Regex(@"((?<!\r)\n|\r\n)", RegexOptions.Compiled);

        public static Int32 CharCount(this String instance, Char c)
        {
            return instance.Length - instance.Replace(c.ToString(), "").Length;
        }

        public static String RemoveSpecialCharacters(this String str)
        {
            StringBuilder sb = new StringBuilder();
            foreach (char c in str)
            {
                if (Char.IsLetterOrDigit(c) || c == '.' || c == '_')
                {
                    sb.Append(c);
                }
            }
            return sb.ToString();
        }

        internal static string EscapeSheetName(this String sheetName)
        {
            if (String.IsNullOrEmpty(sheetName)) return sheetName;

            var needEscape = (!Char.IsLetter(sheetName[0]) && sheetName[0] != '_') ||
                             XLHelper.IsValidA1Address(sheetName) ||
                             XLHelper.IsValidRCAddress(sheetName) ||
                             sheetName.Any(c => (Char.IsPunctuation(c) && c != '.' && c != '_') ||
                                                Char.IsSeparator(c) ||
                                                Char.IsControl(c) ||
                                                Char.IsSymbol(c));
            if (needEscape)
                return String.Concat('\'', sheetName.Replace("'", "''"), '\'');
            else
                return sheetName;
        }

        internal static String FixNewLines(this String value)
        {
            return value.Contains("\n") ? RegexNewLine.Replace(value, Environment.NewLine) : value;
        }

        internal static Boolean PreserveSpaces(this String value)
        {
            return value.StartsWith(" ") || value.EndsWith(" ") || value.Contains(Environment.NewLine);
        }

        internal static String ToCamel(this String value)
        {
            if (value.Length == 0)
                return value;

            if (value.Length == 1)
                return value.ToLower();

            return value.Substring(0, 1).ToLower() + value.Substring(1);
        }

        internal static String ToProper(this String value)
        {
            if (value.Length == 0)
                return value;

            if (value.Length == 1)
                return value.ToUpper();

            return value.Substring(0, 1).ToUpper() + value.Substring(1);
        }

        internal static string UnescapeSheetName(this String sheetName)
        {
            return sheetName
                .Trim('\'')
                .Replace("''", "'");
        }

        internal static string WithoutLast(this String value, int length)
        {
            return length < value.Length ? value.Substring(0, value.Length - length) : String.Empty;
        }

        /// <summary>
        /// Convert a string (containing code units) into code points.
        /// Surrogate pairs of code units are joined to code points.
        /// </summary>
        /// <param name="text">UTF-16 code units to convert.</param>
        /// <param name="output">Output containing code points. Must always be able to fit whole <paramref name="text"/>.</param>
        /// <returns>Number of code points in the <paramref name="output"/>.</returns>
        internal static int ToCodePoints(this ReadOnlySpan<char> text, Span<int> output)
        {
            var j = 0;
            for (var i = 0; i < text.Length; ++i, ++j)
            {
                if (i + 1 < text.Length && char.IsSurrogatePair(text[i], text[i + 1]))
                {
                    output[j] = char.ConvertToUtf32(text[i], text[i + 1]);
                    i++;
                }
                else
                {
                    output[j] = text[i];
                }
            }

            return j;
        }

        /// <summary>
        /// Is the string a new line of any kind (widnows/unix/mac)?
        /// </summary>
        /// <param name="text">Input text to check for EOL at the beginning.</param>
        /// <param name="length">Length of EOL chars.</param>
        /// <returns>True, if text has EOL at the beginning.</returns>
        internal static bool TrySliceNewLine(this ReadOnlySpan<char> text, out int length)
        {
            if (text.Length >= 2 && text[0] == '\r' && text[1] == '\n')
            {
                length = 2;
                return true;
            }

            if (text.Length >= 1 && (text[0] == '\n' || text[0] == '\r'))
            {
                length = 1;
                return true;
            }

            length = default;
            return false;
        }

        /// <summary>
        /// Convert a magic text to a number, where the first letter is in the highest byte of the number.
        /// </summary>
        internal static UInt32 ToMagicNumber(this string magic)
        {
            if (magic.Length > 4)
            {
                throw new ArgumentException();
            }

            return Encoding.ASCII.GetBytes(magic).Select(x => (uint)x).Aggregate((acc, cur) => acc * 256 + cur);
        }

        internal static String TrimFormulaEqual(this String text)
        {
            var trimmed = text.AsSpan().Trim();
            if (trimmed.Length > 1 && trimmed[0] == '=')
                return trimmed[1..].TrimStart().ToString();

            return text;
        }
    }
}
