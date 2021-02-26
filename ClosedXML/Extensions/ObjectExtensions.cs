// Keep this file CodeMaid organised and cleaned
using System;
using System.Globalization;

namespace ClosedXML.Excel
{
    internal static class ObjectExtensions
    {
        public static T CastTo<T>(this Object o)
        {
            return (T)Convert.ChangeType(o, typeof(T));
        }

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

        public static string ToInvariantString<T>(this T value) where T : struct
        {
            return value switch
            {
                sbyte v => v.ToString(CultureInfo.InvariantCulture),
                byte v => v.ToString(CultureInfo.InvariantCulture),
                short v => v.ToString(CultureInfo.InvariantCulture),
                ushort v => v.ToString(CultureInfo.InvariantCulture),
                int v => v.ToString(CultureInfo.InvariantCulture),
                uint v => v.ToString(CultureInfo.InvariantCulture),
                long v => v.ToString(CultureInfo.InvariantCulture),
                ulong v => v.ToString(CultureInfo.InvariantCulture),
                float v => v.ToString("G7", CultureInfo.InvariantCulture), // Specify precision explicitly for backward compatibility
                double v => v.ToString("G15", CultureInfo.InvariantCulture), // Specify precision explicitly for backward compatibility
                decimal v => v.ToString(CultureInfo.InvariantCulture),
                TimeSpan ts => ts.ToString("c", CultureInfo.InvariantCulture),
                DateTime d => d.ToString(CultureInfo.InvariantCulture),
                bool b => b.ToString().ToLowerInvariant(),
                _ => value.ToString(),
            };
        }

        // This method may cause boxing of value types so it is better to replace its calls with
        // the generic version, where applicable
        public static string ObjectToInvariantString(this object value)
        {
            return value switch
            {
                null => string.Empty,
                sbyte v => v.ToString(CultureInfo.InvariantCulture),
                byte v => v.ToString(CultureInfo.InvariantCulture),
                short v => v.ToString(CultureInfo.InvariantCulture),
                ushort v => v.ToString(CultureInfo.InvariantCulture),
                int v => v.ToString(CultureInfo.InvariantCulture),
                uint v => v.ToString(CultureInfo.InvariantCulture),
                long v => v.ToString(CultureInfo.InvariantCulture),
                ulong v => v.ToString(CultureInfo.InvariantCulture),
                float v => v.ToString("G7", CultureInfo.InvariantCulture), // Specify precision explicitly for backward compatibility
                double v => v.ToString("G15", CultureInfo.InvariantCulture), // Specify precision explicitly for backward compatibility
                decimal v => v.ToString(CultureInfo.InvariantCulture),
                TimeSpan ts => ts.ToString("c", CultureInfo.InvariantCulture),
                DateTime d => d.ToString(CultureInfo.InvariantCulture),
                bool b => b.ToString().ToLowerInvariant(),
                _ => value.ToString(),
            };
        }
    }
}
