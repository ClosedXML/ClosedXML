using System;
using System.Globalization;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// A parser of timespan format used by excel during coercion from text to number. <see cref="TimeSpan" /> parsing methods
    /// don't allow for several features required by excel (e.g. seconds/minutes over 60, hours over 24).
    /// Parser can parse following formats from ECMA-376, Part 1, §18.8.30. due to standard text-to-number coercion:
    /// <list type="bullet">
    ///     <item>Format 20 - <c>h:mm</c>.</item>
    ///     <item>Format 21 - <c>h:mm:ss</c>.</item>
    ///     <item>Format 47 - <c>mm:ss.0</c> (format is incorrectly described as <c>mmss.0</c> in the standard,
    ///           but fixed in an implementation errata).</item>
    /// </list>
    /// Timespan is never interpreted through format 45 (<c>mm:ss</c>), instead preferring the format 20 (<c>h:mm</c>).
    /// Timespan is never interpreted through format 46 (<c>[h]:mm:ss</c>], such values are covered by format 21 (<c>h:mm:ss</c>).
    /// </summary>
    /// <remarks>
    /// Note that the decimal fraction differs format 20 and 47, thus mere addition of decimal
    /// place means significantly different values. Parser also copies features of Excel, like whitespaces around
    /// a decimal place (<c>10:20 . 5</c> is allowed).
    /// <example>
    /// <c>20:30</c> is detected as format 20 and the first number is interpreted as hours, thus the serial time is 0.854167.
    /// <c>20:30.0</c> is detected as format 47 and the first number is interpreted as minutes, thus the serial time is 0.014236111.
    /// </example>
    /// </remarks>
    internal static class TimeSpanParser
    {
        public static bool TryParseTime(string s, CultureInfo ci, out TimeSpan result)
        {
            var timeSeparator = ci.DateTimeFormat.TimeSeparator;
            var decimalSeparator = ci.NumberFormat.NumberDecimalSeparator;
            result = default;
            var i = 0;
            SkipWhitespace(ref i, s);
            if (!TryReadNumber(ref i, s, out var hoursOrMinutes))
                return false;

            SkipWhitespace(ref i, s);

            if (!InputMatches(ref i, s, timeSeparator))
                return false;

            SkipWhitespace(ref i, s);
            if (i == s.Length) // Special case ' 10 : '
            {
                result = new TimeSpan(hoursOrMinutes, 0, 0);
                return true;
            }

            if (!TryReadNumber(ref i, s, out var minutesOrSeconds))
                return false;

            SkipWhitespace(ref i, s);

            if (i == s.Length) // Case '10:00'
            {
                result = new TimeSpan(hoursOrMinutes, minutesOrSeconds, 0);
                return hoursOrMinutes < 24 || minutesOrSeconds < 60;
            }

            if (InputMatches(ref i, s, decimalSeparator))
            {
                SkipWhitespace(ref i, s);
                var ms = ReadFractionInMs(ref i, s); // '10:20.' is allowed without digits
                SkipWhitespace(ref i, s);
                result = new TimeSpan(0, 0, hoursOrMinutes, minutesOrSeconds, ms);
                return i == s.Length &&
                       (hoursOrMinutes < 60 || minutesOrSeconds < 60); // No check for min/sec over limit
            }

            // Longer path for h:m:s[.f]
            if (!InputMatches(ref i, s,
                    timeSeparator)) // There is some other character after '10:00', but only ':' ('10:20:0')
                return false;

            SkipWhitespace(ref i, s);
            if (i == s.Length) // Case ' 10 : 0 : '
            {
                result = new TimeSpan(hoursOrMinutes, minutesOrSeconds, 0);
                return hoursOrMinutes < 24 || minutesOrSeconds < 60;
            }

            if (!TryReadNumber(ref i, s, out var seconds)) // Seconds
                return false;

            // At lost two can be over limit
            if ((hoursOrMinutes >= 24 && minutesOrSeconds >= 60)
                || (hoursOrMinutes >= 24 && seconds >= 60)
                || (minutesOrSeconds >= 60 && seconds >= 60))
                return false;


            SkipWhitespace(ref i, s);
            if (i == s.Length) // Case ' 1 : 0 : 0 . '
            {
                result = new TimeSpan(hoursOrMinutes, minutesOrSeconds, seconds);
                return true;
            }

            if (!InputMatches(ref i, s,
                    decimalSeparator)) // The only allowed character is a decimal separator for '1:0:0.'
                return false;

            SkipWhitespace(ref i, s);
            var milliseconds = ReadFractionInMs(ref i, s);
            SkipWhitespace(ref i, s);

            if (i == s.Length) // Case ' 1 : 0 : 0 . 0 '
            {
                result = new TimeSpan(0, hoursOrMinutes, minutesOrSeconds, seconds, milliseconds);
                return (hoursOrMinutes < 24 && minutesOrSeconds < 60)
                       || (hoursOrMinutes < 24 && seconds < 60)
                       || (minutesOrSeconds < 60 && seconds < 60); // Just one 1 field under limit is enough
            }

            return false; // There was some unexpected chars at the end

            static bool TryReadNumber(ref int i, string t, out int num)
            {
                var start = i;
                num = 0;
                var digitCount = 0;
                while (i < t.Length && t[i] >= '0' && t[i] <= '9')
                {
                    num = num * 10 + t[i] - '0';
                    digitCount++;
                    i++;
                }

                if (digitCount == 0 || num > 9999)
                    return false;
                if (t[start] == '0' && digitCount > 2)
                    return false;
                return true;
            }

            static int ReadFractionInMs(ref int i, string t)
            {
                var num = 0;
                var digitCount = 0;
                while (i < t.Length && t[i] >= '0' && t[i] <= '9')
                {
                    num = num * 10 + t[i] - '0';
                    digitCount++;
                    i++;
                }

                // Maximum resolution is 1 ms of pattern
                return (int)Math.Round(num / Math.Pow(10, digitCount - 3), MidpointRounding.AwayFromZero);
            }

            static void SkipWhitespace(ref int i, string t)
            {
                while (i < t.Length && t[i] == ' ') i++;
            }

            static bool InputMatches(ref int i, string t, string expected)
            {
                for (var expectedIdx = 0; expectedIdx < expected.Length; ++expectedIdx)
                {
                    var inputIdx = i + expectedIdx;
                    if (inputIdx == t.Length || expected[expectedIdx] != t[inputIdx])
                    {
                        return false;
                    }
                }

                i += expected.Length; // Branch can differ depending on the input (':' vs '.'), so move only when input matches
                return true;
            }
        }
    }
}
