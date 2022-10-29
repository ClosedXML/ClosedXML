using System;
using System.Collections.Concurrent;
using System.Globalization;
using System.Linq;

namespace ClosedXML.Excel.CalcEngine
{
    internal static class DateTimeParser
    {
        private const DateTimeStyles Style = DateTimeStyles.NoCurrentDateDefault | DateTimeStyles.AllowInnerWhite | DateTimeStyles.AllowTrailingWhite;

        // It's highly likely that Excel has its own database of culture specific patterns for parsing.
        // Excel has it's own parser (that accepts 1900-02-29 ^_^), never seems to parse name of a day,
        // values of hours can be up to 9999 and safely overflow...
        // Although for displaying, Excel takes a cue from region setting pattern, not so for parsing (at least
        // couldn't produce observable difference by changing setting of a culture in region dialogue).
        // .NET Core and .NET Framework also produce different patterns for GetAllDateTimePatterns.
        // This is not a perfect solution by any means, but best we can do in absence of knowledge
        // what patterns Excel uses for which cultures.
        private static readonly ConcurrentDictionary<CultureInfo, string[]> CultureSpecificPatterns = new();

        private static readonly string[] TimeOfDayPatterns = { "h:m tt", "h:m t", "h:m:s tt", "h:m:s t" };

        public static bool TryParseCultureDate(string s, CultureInfo culture, out DateTime date)
        {
            var datePatterns = CultureSpecificPatterns.GetOrAdd(culture, static ci =>
            {
                var shortDatePatterns = ci.DateTimeFormat.GetAllDateTimePatterns('d')
                    .Concat(ci.DateTimeFormat.GetAllDateTimePatterns('D'))
                    .Where(pattern => !pattern.Contains("dddd")) // It doesn't seem that Excel parser is capable of parsing day names in any culture
                    .Distinct().ToArray();

                // Not sure about this, but reasonably close. Hours pattern is probably generated (e.g. 'as-IN' culture
                // has AM designator before hours in patterns, but Excel requires it to be at the end). There most likely
                // isn't a pattern to just use. Example: for en-US, Excel type coercion can transform "aug 10, 2022 14:10",
                // but every single format from CultureInfo.DateTimeFormat requires AM/PM. and two digits for minutes (thus
                // the input couldn't match in any format => excel has likely it's own logic, independent of region setting).
                var timePatterns = new[] { "h:m tt", "H:m", "h:m" };
                var longDatePatterns = shortDatePatterns
                    .SelectMany(datePattern => timePatterns.Select(timePattern => FormattableString.Invariant($"{datePattern} {timePattern}")));

                // ISO8601 should be parseable in all cultures, not sure if Excel does.
                return shortDatePatterns.Concat(longDatePatterns).Concat(new[] { "yyyy-MM-DD" }).Distinct().ToArray();
            });

            return DateTime.TryParseExact(s, datePatterns, culture, Style, out date);
        }

        public static bool TryParseTimeOfDay(string s, CultureInfo c, out DateTime timeOfDay)
        {
            if (DateTime.TryParseExact(s, TimeOfDayPatterns, c, Style, out timeOfDay))
                return true;

            if (DateTime.TryParseExact(s, TimeOfDayPatterns, CultureInfo.InvariantCulture, Style, out timeOfDay))
                return true;

            return false;
        }
    }
}
