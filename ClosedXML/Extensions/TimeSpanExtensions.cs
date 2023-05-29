#nullable disable

using ClosedXML.Excel.CalcEngine;
using System;
using System.Globalization;
using System.Text;

namespace ClosedXML.Excel
{
    internal static class TimeSpanExtensions
    {
        public static double ToSerialDateTime(this TimeSpan time)
        {
            return time.Ticks / (double)TimeSpan.TicksPerDay;
        }

        /// <summary>
        /// Return a string representation of a TimeSpan that can be parsed by an Excel through text-to-number coercion.
        /// </summary>
        /// <remarks>
        /// Excel can convert time span string back to a number, but only if it doesn't has days in the string, only hours.
        /// It's an opposite of <see cref="TimeSpanParser"/>.
        /// </remarks>
        public static String ToExcelString(this TimeSpan ts, CultureInfo culture)
        {
            var timeSep = culture.DateTimeFormat.TimeSeparator;
            var sb = new StringBuilder()
                .Append(ts.Hours + 24 * ts.Days).Append(timeSep)
                .AppendFormat("{0:D2}", ts.Minutes).Append(timeSep)
                .AppendFormat("{0:D2}", ts.Seconds);
            // the ts.Miliseconds property uses whole division and due to serial datetime conversion, it should be rounded instead
            var ms = (int)Math.Round((ts.Ticks % TimeSpan.TicksPerSecond) * 1000.0 / (TimeSpan.TicksPerSecond));
            if (ms != 0)
            {
                sb.Append(culture.NumberFormat.CurrencyDecimalSeparator);
                if (ms % 100 == 0)
                    sb.AppendFormat("{0:0}", ms / 100);
                else if (ms % 10 == 0)
                    sb.AppendFormat("{0:00}", ms / 10);
                else
                    sb.AppendFormat("{0:000}", ms);
            }
            return sb.ToString();
        }
    }
}
