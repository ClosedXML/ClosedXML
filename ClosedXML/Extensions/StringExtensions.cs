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
    }
}
