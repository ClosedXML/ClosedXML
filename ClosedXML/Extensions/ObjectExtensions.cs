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
                    // Specify precision explicitly for backward compatibility
                    return v.ToString("G7", CultureInfo.InvariantCulture);

                case double v:
                    // Specify precision explicitly for backward compatibility
                    return v.ToString("G15", CultureInfo.InvariantCulture);

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

        // This method may cause boxing of value types so it is better to replace its calls with
        // the generic version, where applicable
        public static string ObjectToInvariantString(this object value)
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
                    // Specify precision explicitly for backward compatibility
                    return v.ToString("G7", CultureInfo.InvariantCulture);

                case double v:
                    // Specify precision explicitly for backward compatibility
                    return v.ToString("G15", CultureInfo.InvariantCulture);

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
