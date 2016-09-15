﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using System.IO;
using System.Xml;
using ClosedXML.Utils;
using DocumentFormat.OpenXml;

[assembly: CLSCompliantAttribute(true)]
namespace ClosedXML.Excel
{
    public static class Extensions
    {
        // Adds the .ForEach method to all IEnumerables


        private static readonly char[] hexDigits = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F'};

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
                if (distinctItems.Contains(item))
                    return true;
                else
                    distinctItems.Add(item);
            }
            return false;
        }

        public static T CastTo<T>(this Object o)
        {
            return (T)Convert.ChangeType(o, typeof(T));
        }

    }

    public static class DictionaryExtensions
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

    public static class StringExtensions
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

    public static class DateTimeExtensions
    {
        public static Double MaxOADate
        {
            get
            {
                return 2958465.99999999;
            }
        }

        public static DateTime NextWorkday(this DateTime date, List<DateTime> bankHolidays)
        {
            var nextDate = date.AddDays(1);
            while (nextDate.DayOfWeek == DayOfWeek.Saturday
                || nextDate.DayOfWeek == DayOfWeek.Sunday
                || bankHolidays.Contains(nextDate)
                )
                nextDate = nextDate.AddDays(1);

            return nextDate;
        }
    }

    public static class IntegerExtensions
    {
        private static readonly NumberFormatInfo nfi = CultureInfo.InvariantCulture.NumberFormat;
        [ThreadStatic]
        private static Dictionary<Int32, String> intToString;
        public static String ToInvariantString(this Int32 value)
        {
            String sValue;
            if (intToString == null)
            {
                intToString = new Dictionary<int, string>();
                sValue = value.ToString(nfi);
                intToString.Add(value, sValue);
            }
            else
            {
                if (!intToString.TryGetValue(value, out sValue))
                {
                    sValue = value.ToString(nfi);
                    intToString.Add(value, sValue);
                }
            }
            return sValue;
        }

    }

    public static class DoubleExtensions
    {
        //All numbers are stored in XL files as invarient culture this is just a easy helper
        public static String ToInvariantString(this Double value)
        {
            return value.ToString(CultureInfo.InvariantCulture);
        }

        public static Double SaveRound(this Double value)
        {
            return Math.Round(value, 6);
        }
    }

    public static class FontBaseExtensions
    {

        private static Font GetCachedFont(IXLFontBase fontBase, Dictionary<IXLFontBase, Font> fontCache)
        {
            Font font;
            if (!fontCache.TryGetValue(fontBase, out font))
            {
                font = new Font(fontBase.FontName, (float)fontBase.FontSize, GetFontStyle(fontBase));
                fontCache.Add(fontBase, font);
            }
            return font;
        }

        public static Double GetWidth(this IXLFontBase fontBase, String text, Dictionary<IXLFontBase, Font> fontCache)
        {
            if (XLHelper.IsNullOrWhiteSpace(text))
                return 0;

            var font = GetCachedFont(fontBase, fontCache);

            var textSize = GraphicsUtils.MeasureString(text, font);

            double width = (((textSize.Width / (double)7) * 256) - (128 / 7)) / 256;
            width = (double)decimal.Round((decimal)width + 0.2M, 2);

            return width;
        }

        private static FontStyle GetFontStyle(IXLFontBase font)
        {
            FontStyle fontStyle = FontStyle.Regular;
            if (font.Bold) fontStyle |= FontStyle.Bold;
            if (font.Italic) fontStyle |= FontStyle.Italic;
            if (font.Strikethrough) fontStyle |= FontStyle.Strikeout;
            if (font.Underline != XLFontUnderlineValues.None ) fontStyle |= FontStyle.Underline;
            return fontStyle;
        }

        public static Double GetHeight(this IXLFontBase fontBase, Dictionary<IXLFontBase, Font> fontCache)
        {
            var font = GetCachedFont(fontBase, fontCache);
            var textSize = GraphicsUtils.MeasureString("X", font);
            return (double)textSize.Height * 0.85;
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
        }
    }

    public static class XDocumentExtensions
    {
        public static XDocument Load(Stream stream)
        {
            using (XmlReader reader = XmlReader.Create(stream))
            {
                return XDocument.Load(reader);
            }
        }
    }

    public static class EnumerableExtensions
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

    public static class DoubleValueExtensions
    {
        public static DoubleValue SaveRound(this DoubleValue value)
        {
            return value.HasValue ? new DoubleValue(Math.Round(value, 6)) : value;
        }
    }
}

