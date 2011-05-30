using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Globalization;


[assembly: CLSCompliantAttribute(true)]
namespace ClosedXML
{
    public static class Extensions
    {
        // Adds the .ForEach method to all IEnumerables
        public static void ForEach<T>(this IEnumerable<T> source, Action<T> action)
        {
            foreach (T item in source)
                action(item);
        }

        static char[] hexDigits = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F'};

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
                //if ((c >= '0' && c <= '9') || (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') | c == '.' || c == '_')
                if (Char.IsLetterOrDigit(c))
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
        public static bool IsNullOrWhiteSpace(string value)
        {
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
    }

    public static class IntegerExtensions
    {
        private static NumberFormatInfo nfi = CultureInfo.InvariantCulture.NumberFormat;
        private static Dictionary<Int32, String> intToString = new Dictionary<int, string>();
        public static String ToStringLookup(this Int32 value)
        {
            if (!intToString.ContainsKey(value))
            {
                intToString.Add(value, value.ToString(nfi));
            }
            return intToString[value];
        }
    }


}

