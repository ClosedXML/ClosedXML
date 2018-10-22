using ClosedXML.Utils;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

[assembly: CLSCompliantAttribute(true)]

namespace ClosedXML.Excel
{
    internal static class Extensions
    {
        // Adds the .ForEach method to all IEnumerables

        private static readonly char[] hexDigits = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F' };

        public static String ToHex(this Color color)
        {
            byte[] bytes = new byte[4];

            bytes[0] = color.A;

            bytes[1] = color.R;

            bytes[2] = color.G;

            bytes[3] = color.B;

            char[] chars = new char[bytes.Length * 2];

            for (int i = 0; i < bytes.Length; i++)
            {
                int b = bytes[i];

                chars[i * 2] = hexDigits[b >> 4];

                chars[i * 2 + 1] = hexDigits[b & 0xF];
            }

            return new string(chars);
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

        public static Int32 CharCount(this String instance, Char c)
        {
            return instance.Length - instance.Replace(c.ToString(), "").Length;
        }

        public static Boolean HasDuplicates<T>(this IEnumerable<T> source)
        {
            HashSet<T> distinctItems = new HashSet<T>();
            foreach (var item in source)
            {
                if (!distinctItems.Add(item))
                {
                    return true;
                }
            }
            return false;
        }

        public static T CastTo<T>(this Object o)
        {
            return (T)Convert.ChangeType(o, typeof(T));
        }
    }

    internal static class DictionaryExtensions
    {
        public static void RemoveAll<TKey, TValue>(this Dictionary<TKey, TValue> dic,
            Func<TValue, bool> predicate)
        {
            var keys = dic.Keys.Where(k => predicate(dic[k])).ToList();
            foreach (var key in keys)
            {
                dic.Remove(key);
            }
        }
    }

    internal static class StringExtensions
    {
        private static readonly Regex RegexNewLine = new Regex(@"((?<!\r)\n|\r\n)", RegexOptions.Compiled);

        public static String FixNewLines(this String value)
        {
            return value.Contains("\n") ? RegexNewLine.Replace(value, Environment.NewLine) : value;
        }

        public static Boolean PreserveSpaces(this String value)
        {
            return value.StartsWith(" ") || value.EndsWith(" ") || value.Contains(Environment.NewLine);
        }

        public static String ToCamel(this String value)
        {
            if (value.Length == 0)
                return value;

            if (value.Length == 1)
                return value.ToLower();

            return value.Substring(0, 1).ToLower() + value.Substring(1);
        }

        public static String ToProper(this String value)
        {
            if (value.Length == 0)
                return value;

            if (value.Length == 1)
                return value.ToUpper();

            return value.Substring(0, 1).ToUpper() + value.Substring(1);
        }
    }

    internal static class DateTimeExtensions
    {
        public static Double MaxOADate
        {
            get
            {
                return 2958465.99999999;
            }
        }

        public static DateTime NextWorkday(this DateTime date, IEnumerable<DateTime> bankHolidays)
        {
            var nextDate = date.AddDays(1);
            while (!nextDate.IsWorkDay(bankHolidays))
                nextDate = nextDate.AddDays(1);

            return nextDate;
        }

        public static DateTime PreviousWorkDay(this DateTime date, IEnumerable<DateTime> bankHolidays)
        {
            var previousDate = date.AddDays(-1);
            while (!previousDate.IsWorkDay(bankHolidays))
                previousDate = previousDate.AddDays(-1);

            return previousDate;
        }

        public static bool IsWorkDay(this DateTime date, IEnumerable<DateTime> bankHolidays)
        {
            return date.DayOfWeek != DayOfWeek.Saturday
                && date.DayOfWeek != DayOfWeek.Sunday
                && !bankHolidays.Contains(date);
        }
    }

    internal static class IntegerExtensions
    {
        public static bool Between(this int val, int from, int to)
        {
            return val >= from && val <= to;
        }
    }

    internal static class DecimalExtensions
    {
        public static Decimal SaveRound(this Decimal value)
        {
            return Math.Round(value, 6);
        }
    }

    internal static class DoubleExtensions
    {
        public static Double SaveRound(this Double value)
        {
            return Math.Round(value, 6);
        }
    }

    internal static class FontBaseExtensions
    {
        public static Double GetWidth(this IXLFontBase fontBase, String text, Dictionary<IXLFontBase, Font> fontCache)
        {
            if (String.IsNullOrWhiteSpace(text))
                return 0;

            var font = GetCachedFont(fontBase, fontCache);
            var textWidth = GraphicsUtils.MeasureString(text, font).Width;

            double width = (textWidth / 7d * 256 - 128 / 7) / 256;
            width = Math.Round(width + 0.2, 2);

            return width;
        }

        public static Double GetHeight(this IXLFontBase fontBase, Dictionary<IXLFontBase, Font> fontCache)
        {
            var font = GetCachedFont(fontBase, fontCache);
            var textHeight = GraphicsUtils.MeasureString("X", font).Height;
            return (double)textHeight * 0.85;
        }

        public static void CopyFont(this IXLFontBase font, IXLFontBase sourceFont)
        {
            font.Bold = sourceFont.Bold;
            font.Italic = sourceFont.Italic;
            font.Underline = sourceFont.Underline;
            font.Strikethrough = sourceFont.Strikethrough;
            font.VerticalAlignment = sourceFont.VerticalAlignment;
            font.Shadow = sourceFont.Shadow;
            font.FontSize = sourceFont.FontSize;
            font.FontColor = sourceFont.FontColor;
            font.FontName = sourceFont.FontName;
            font.FontFamilyNumbering = sourceFont.FontFamilyNumbering;
            font.FontCharSet = sourceFont.FontCharSet;
        }

        private static Font GetCachedFont(IXLFontBase fontBase, Dictionary<IXLFontBase, Font> fontCache)
        {
            if (!fontCache.TryGetValue(fontBase, out Font font))
            {
                font = new Font(fontBase.FontName, (float)fontBase.FontSize, GetFontStyle(fontBase));
                fontCache.Add(fontBase, font);
            }
            return font;
        }

        private static FontStyle GetFontStyle(IXLFontBase font)
        {
            FontStyle fontStyle = FontStyle.Regular;
            if (font.Bold) fontStyle |= FontStyle.Bold;
            if (font.Italic) fontStyle |= FontStyle.Italic;
            if (font.Strikethrough) fontStyle |= FontStyle.Strikeout;
            if (font.Underline != XLFontUnderlineValues.None) fontStyle |= FontStyle.Underline;
            return fontStyle;
        }
    }

    internal static class XDocumentExtensions
    {
        public static XDocument Load(Stream stream)
        {
            using (XmlReader reader = XmlReader.Create(stream))
            {
                return XDocument.Load(reader);
            }
        }
    }

    internal static class EnumerableExtensions
    {
        public static void ForEach<T>(this IEnumerable<T> source, Action<T> action)
        {
            foreach (T item in source)
                action(item);
        }

        public static Type GetItemType<T>(this IEnumerable<T> source)
        {
            return typeof(T);
        }
    }

    internal static class ListExtensions
    {
        public static void RemoveAll<T>(this IList<T> list, Func<T, bool> predicate)
        {
            var indices = list.Where(item => predicate(item)).Select((item, i) => i).OrderByDescending(i => i).ToList();
            foreach (var i in indices)
            {
                list.RemoveAt(i);
            }
        }
    }

    internal static class DoubleValueExtensions
    {
        public static DoubleValue SaveRound(this DoubleValue value)
        {
            return value.HasValue ? new DoubleValue(Math.Round(value, 6)) : value;
        }
    }

    internal static class TypeExtensions
    {
        public static bool IsNumber(this Type type)
        {
            return type == typeof(sbyte)
                    || type == typeof(byte)
                    || type == typeof(short)
                    || type == typeof(ushort)
                    || type == typeof(int)
                    || type == typeof(uint)
                    || type == typeof(long)
                    || type == typeof(ulong)
                    || type == typeof(float)
                    || type == typeof(double)
                    || type == typeof(decimal);
        }

        public static bool IsSimpleType(this Type type)
        {
            return type.IsPrimitive
                || type == typeof(String)
                || type == typeof(DateTime)
                || type == typeof(TimeSpan)
                || type.IsNumber();
        }
    }

    internal static class ObjectExtensions
    {
        public static bool IsNumber(this object value)
        {
            return value is sbyte
                    || value is byte
                    || value is short
                    || value is ushort
                    || value is int
                    || value is uint
                    || value is long
                    || value is ulong
                    || value is float
                    || value is double
                    || value is decimal;
        }

        public static string ToInvariantString(this object value)
        {
            if (value == null)
                return string.Empty;

            switch (value)
            {
                case sbyte v:
                    return v.ToString(CultureInfo.InvariantCulture);

                case byte v:
                    return v.ToString(CultureInfo.InvariantCulture);

                case short v:
                    return v.ToString(CultureInfo.InvariantCulture);

                case ushort v:
                    return v.ToString(CultureInfo.InvariantCulture);

                case int v:
                    return v.ToString(CultureInfo.InvariantCulture);

                case uint v:
                    return v.ToString(CultureInfo.InvariantCulture);

                case long v:
                    return v.ToString(CultureInfo.InvariantCulture);

                case ulong v:
                    return v.ToString(CultureInfo.InvariantCulture);

                case float v:
                    return v.ToString(CultureInfo.InvariantCulture);

                case double v:
                    return v.ToString(CultureInfo.InvariantCulture);

                case decimal v:
                    return v.ToString(CultureInfo.InvariantCulture);

                case TimeSpan ts:
                    return ts.ToString("c", CultureInfo.InvariantCulture);

                case DateTime d:
                    return d.ToString(CultureInfo.InvariantCulture);

                default:
                    return value.ToString();
            }
        }
    }
}
