// Keep this file CodeMaid organised and cleaned
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ClosedXML.Excel
{
    internal static class StringExtensions
    {
        private static readonly Regex RegexNewLine = new Regex(@"((?<!\r)\n|\r\n)", RegexOptions.Compiled);

        public static int CharCount(this string instance, char c)
        {
            return instance.Length - instance.Replace(c.ToString(), "").Length;
        }

        public static string RemoveSpecialCharacters(this string str)
        {
            var sb = new StringBuilder();
            foreach (var c in str)
            {
                if (char.IsLetterOrDigit(c) || c == '.' || c == '_')
                {
                    sb.Append(c);
                }
            }
            return sb.ToString();
        }

        internal static string EscapeSheetName(this string sheetName)
        {
            if (string.IsNullOrEmpty(sheetName))
            {
                return sheetName;
            }

            var needEscape = (!char.IsLetter(sheetName[0]) && sheetName[0] != '_') ||
                             XLHelper.IsValidA1Address(sheetName) ||
                             XLHelper.IsValidRCAddress(sheetName) ||
                             sheetName.Any(c => (char.IsPunctuation(c) && c != '.' && c != '_') ||
                                                char.IsSeparator(c) ||
                                                char.IsControl(c) ||
                                                char.IsSymbol(c));
            if (needEscape)
            {
                return string.Concat('\'', sheetName.Replace("'", "''"), '\'');
            }
            else
            {
                return sheetName;
            }
        }

        internal static string FixNewLines(this string value)
        {
            return value.Contains("\n") ? RegexNewLine.Replace(value, XLConstants.NewLine) : value;
        }

        internal static bool PreserveSpaces(this string value)
        {
            return value.StartsWith(" ") || value.EndsWith(" ") || value.Contains(XLConstants.NewLine);
        }

        internal static string ToCamel(this string value)
        {
            if (value.Length == 0)
            {
                return value;
            }

            if (value.Length == 1)
            {
                return value.ToLower();
            }

            return value.Substring(0, 1).ToLower() + value.Substring(1);
        }

        internal static string ToProper(this string value)
        {
            if (value.Length == 0)
            {
                return value;
            }

            if (value.Length == 1)
            {
                return value.ToUpper();
            }

            return value.Substring(0, 1).ToUpper() + value.Substring(1);
        }

        internal static string UnescapeSheetName(this string sheetName)
        {
            return sheetName
                .Trim('\'')
                .Replace("''", "'");
        }
    }
}