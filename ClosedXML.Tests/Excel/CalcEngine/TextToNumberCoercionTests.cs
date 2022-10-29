using System;
using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    [SetCulture("en-US")]
    public class TextToNumberCoercionTests
    {
        private const double Tolerance = 0.000001;

        [Test]
        public void TimeSpan_MaximumResolutionIsOneMs()
        {
            var firstValue = (double)XLWorkbook.EvaluateExpr("\"0:0:0.0015\" * 1");
            var secondValue = (double)XLWorkbook.EvaluateExpr("\"0:0:0.0024\" * 1");
            Assert.AreEqual(firstValue, secondValue);
        }

        [TestCase("100%", 1)]
        [TestCase("-100%", -1)]
        [TestCase("200%", 2)]
        [TestCase("0000%", 0)]
        [TestCase("1%", 0.01)]
        [TestCase("+1%", 0.01)]
        [TestCase(" -75 % ", -0.75)]
        [TestCase(" - 100 % ", -1, Ignore = ".NET parser doesn't allow whitespace between sign and number.")]
        public void Percent_Format9(string percent, double? expectedValue) // Format 9 '0%'
        {
            AssertCoercion(percent, expectedValue);
        }

        [TestCase("100.5%", 1.005)]
        [TestCase("100 . 5%", null)]
        [TestCase(" - 100.59 % ", -1.0059, Ignore = ".NET parser doesn't allow whitespace between sign and number.")]
        [TestCase("0.123456%", 0.00123456)]
        [TestCase(".5%", 0.005)]
        [TestCase("  -.375 % ", -0.00375)]
        [TestCase("100.%", 1)]
        public void Percent_Format10(string percent, double? expectedValue) // Format 10 '0.00%'
        {
            AssertCoercion(percent, expectedValue);
        }

        [TestCase("(100%)", -1, Ignore = ".NET parser doesn't parse percents.")]
        [TestCase("(-100%)", null)] // Can't have minus sign inside the brackets
        [TestCase("-(100%)", null)] // Can't have minus sign outside the brackets
        [TestCase("1,000.00%", 10)]
        [TestCase("(1,000.00%)", -10, Ignore = ".NET parser doesn't parse percents.")]
        [TestCase(" % 100", 1)] // Percents can be at start or end, position doesn't matter
        public void Percent_UnlistedFormats(string percent, double? expectedValue) // 
        {
            AssertCoercion(percent, expectedValue);
        }

        [TestCase("0 1/2", 0.5)]
        [TestCase("0 /20", null)]
        [TestCase("0 1/32768", null)] // Denominator can be at most 2^15-1
        [TestCase("0 1/32767", 3.0518509475997192E-05d)]
        [TestCase("0 32768/1", null)] // Nominator can be at most 2^15-1
        [TestCase("0 32767/1", 32767)]
        [TestCase("1 32767/032767", null)] // Fraction can be only 5 digits at most
        [TestCase("1 00100/025", 5)]
        [TestCase("1 100/-2", null)] // Fractions can't be negative
        [TestCase("1 -1/2", null)]
        [TestCase("- 1 1/2", -1.5)] // can use minus sign
        [TestCase("+1 1/2", 1.5)] // or plus sign
        [TestCase("1.5 1/2", null)] // Can't use dot in whole part
        [TestCase("   1 10/20  ", 1.5)]
        [TestCase("1  1/2", null)] // Between whole part and nominator must be exactly one space
        [TestCase("1 1 /2", null)] // Can't have spaces between nominator and denominator
        [TestCase("1 1/ 2", null)]
        [TestCase("1	1/2", null)] // Tab and other whitespaces aren't allowed
        [TestCase("0 1/0", null)] // Division by zero
        public void Fraction_Format12_13(string fraction, double? expectedValue) // Format 12+13 '# ??/??' and  '# ?/?'
        {
            AssertCoercion(fraction, expectedValue);
        }

        [TestCase("02/28/20", 43889)]
        [TestCase("002/28/20", null)]
        [TestCase("02/028/20", null)]
        [TestCase("02/28/022", null)]
        public void Date_Format14(string date, double? expectedValue) // Format 14 is taken from region setting, but for en (and MS errata) says 'm/d/yyyy'
        {
            AssertCoercion(date, expectedValue);
        }

        [TestCase("30-apr-2000", 36646)]
        [TestCase("30-apr-20", 43951)] // 2020-04-30	
        [TestCase("31-dec-9999", 2958465)]
        [TestCase("1-jan-10000", null)]
        [TestCase("1 - jan - 2022  ", 44562)] // Can have whitespace in the date
        [TestCase(" 1-jan-2022", null, Ignore = ".NET parser doesn't respect the whitespace styles of a date during parsing.")] // Can't have whitespaces at the start
        [TestCase("31-dec-1899", null)] // Check 1900 "leap" year issue...
        [TestCase("1-jan-1900", 1)]
        [TestCase("28-feb-1900", 59)]
        [TestCase("1-mar-1900", 61)]
        public void Date_Format15(string date, double? expectedValue) // Format 15 d-mmm-yy
        {
            AssertCoercion(date, expectedValue);
        }

        [TestCase("0-mar", null)] // Zero day not accepted
        [TestCase("1-mar", 44621)]
        [TestCase("1-marc", 44621, Ignore = ".NET parser recognizes only abbreviation or full name of a month.")]
        [TestCase("1-march", 44621)]
        [TestCase(" 1 - apr  ", 44652)] // Unlike many others, this format also allows space at the start, not just inside and at the end
        [TestCase("31-apr", null)] // April has only 30 days
        public void Date_Format16(string text, double? expectedValue) // Format 16 'd-mmm'
        {
            if (expectedValue is not null)
            {
                var date = DateTime.FromOADate(expectedValue.Value);
                expectedValue = new DateTime(DateTime.Now.Year, date.Month, date.Day).ToOADate();
            }

            AssertCoercion(text, expectedValue);
        }

        // In en locale, there should be an extra pattern MMM-dd that is before the standard MMM-yy, but .NET Framework doesn't have it.
        // To overcome missing locale, use numbers over 31 for year (otherwise they should be interpreted as days)
        [TestCase("jan-02", 44563, Ignore = ".NET misses culture, en interprets it as MMM-dd, but czech as MMM-yy, so the MMM-dd is the extra culture for en.")] // interpreted as 2022-01-02
        [TestCase("jan-31", 44592, Ignore = "Missing excel culture mapping")] // 2022-01-02
        [TestCase("jan-32", 11689)] // 1932-01-01
        [TestCase("feb-29", 47150, Ignore = "Missing excel culture mapping")] // 2029-02-01
        [TestCase("feb-30", 10990)] // 1930-02-01
        [TestCase("feb-31", 11355)] // 1931-02-01
        [TestCase("feb-003", null)] // three digits not allowed
        [TestCase("aug   -   55", 20302)] // spaces are allowed inside the pattern
        [TestCase(" aug-55", null, Ignore = ".NET allow whitespaces even without specified DateTimeStyle.AllowLeadingWhite")] // starting spaces not allowed 
        [TestCase("aug-55 ", 20302)] // trailing spaces allowed
        [TestCase("MaR-42", 15401)] // case insensitive
        [TestCase("marc-2", 44622, Ignore = ".NET parser recognizes only abbreviation or full name of a month.")] // name can be more than three long abbr
        [TestCase("march-55", 20149)]
        [TestCase("ma-2", null)] // Name of month must be at least three chars long	
        public void Date_Format17(string text, double? expectedValue) // Format 17 'mmm-yy'
        {
            AssertCoercion(text, expectedValue);
        }

        [TestCase("1:20 AM", 0.055555555555555552d)]
        [TestCase("1:20 aM", 0.055555555555555552d)]
        [TestCase("1:60 AM", null)] // Minutes must be 0-59 range
        [TestCase("1:59 AM", 0.082638888888888887d)]
        [TestCase("13:00 AM", null)] // AM only allows hours in 0-12 range
        [TestCase("7:30 A", 0.3125)] // only starting letter of AM
        [TestCase("1:9 AM", 0.04791666666666667d)] // Single digit minutes
        public void Date_Format18(string text, double? expectedValue) // Format 18 'h:mm AM/PM'
        {
            AssertCoercion(text, expectedValue);
        }

        [TestCase("12:0:0 PM", 0.5)]
        [TestCase("12:0:18 aM", 0.00020833333333333335d)] // case insensitive AM designator
        [TestCase("13:0:0 PM", null)] // hours can't be outside of 0-12, unlike other format
        [TestCase("13:0:0 AM", null)]
        [TestCase("00:60:00 AM", null)] // minutes can't be outside of 0-59, unlike other format
        [TestCase("00:59:00 AM", 0.040972222222222222d)]
        [TestCase("00:00:60 AM", null)] // seconds can't be outside of 0-59, unlike other format
        [TestCase("00:00:59 AM", 0.00068287037037037036d)]
        [TestCase("00:00: AM", null)] // can't omit second part (differs from time span).
        [TestCase("1:2:3 AM", 0.043090277777777776d)]
        public void Date_Format19(string text, double? expectedValue) // Format 19 'h:mm:ss AM/PM'
        {
            AssertCoercion(text, expectedValue);
        }

        [TestCase("2/5/2022 0:0", 44597)]
        [TestCase("05/5/2022 0:0", 44686)] // Extra zero padding allowed
        [TestCase("005/5/2022 0:0", null)] // 0 prefix requires at most 2 digits
        [TestCase("13/5/2022 0:0", null)] // Month outside of range
        [TestCase("11/030/2022 0:0", null)]
        [TestCase("11/30/02022 0:0", null)] // Extra zero before year not allowed
        [TestCase("11/30/2022 24:59", 44896.04097, Ignore = "Excel can have out of range parts, but .NET parsers can't.")]
        [TestCase("11/30/2022 24:60", null)] // Both parts are out of range
        [TestCase("11/30/2022 23:160", 44896.06944, Ignore = "Excel can have one of of range part, but .NET parser can't.")]
        [TestCase("11/30/2022 9999:59", 45311.66597, Ignore = "Excel parser accepts numbers over limit for hours.")]
        [TestCase("11/30/2022 10000:59", null)] // Hours can't be over 9999
        [TestCase("aug 10, 2022 14:10", 44783.590277777781d, Ignore = "Excel specific parsing of months accepts anything from three letters up to full name, but such pattern is not in any en-US DateTimeFormat pattern.")]
        [TestCase("august 10, 2022 14:10", 44783.590277777781d)]
        public void DateTime_Format22(string text, double? expectedValue) // Format 22 'm/d/yyyy h:mm'. Specification incorrectly states 'm/d/yy h:mm', but fixed per MS errata.
        {
            AssertCoercion(text, expectedValue);
        }

        [TestCase("00:00", 0)] // Can parse zero
        [TestCase("90:00", 3.75)] // Minutes can be can be over 60
        [TestCase("59:59", 2.499305556)] // Even if looks like mm:ss, it is actually parsed as h:mm
        [TestCase("10:", 0.416666667)] // Last part can be omitted and zero is used
        [TestCase("9999:", 416.625)] // Upper limit of first part is parseable
        [TestCase("10000:", null)] // Part value over a limit is not parseable
        [TestCase(":5", null)] // Can't omit first part
        [TestCase("24:60", null)] // Only one part can be outside of limit, here are both
        [TestCase("30:59", 1.290972222)] // Hour part can be over 23
        [TestCase("23:300", 1.166666667)] // Minute part over over 59
        public void TimeSpan_Format20(string timeSpan, double? expectedValue) // 'h:mm'
        {
            AssertCoercion(timeSpan, expectedValue, Tolerance);
        }

        [TestCase("0:01:01", 0.000706019)]
        [TestCase("000:01:01", null)] // Extra zeros.
        [TestCase("00:001:01", null)] // Three digits in a part that starts with 0
        [TestCase("0:01:001", null)] // Three digits in a part that starts with 0
        [TestCase("00:60:60", null)] // Only one part can be over the limit, but here are minutes and seconds
        [TestCase("24:60:00", null)] // Only one part can be over the limit, but here are hours and minutes
        [TestCase("24:00:60", null)] // Only one part can be over the limit, but here are hours and seconds
        [TestCase("23:60:06", 1.000069444)]
        [TestCase("  24   :  00  :   59  ", 1.00068287)] // Extra padding
        [TestCase("24:0:", 1)] // Last part can be omitted 
        [TestCase("0::0", null)] // Parts in the middle can't be omitted
        [TestCase(":0:0", null)] // First part can't be omitted
        public void TimeSpan_Format21(string timeSpan, double? expectedValue) // 'h:mm:ss'
        {
            AssertCoercion(timeSpan, expectedValue, Tolerance);
        }

        [TestCase("14:30.0", 0.010069444)] // Happy case, can be over 12 (to differ from AM/PM times)
        [TestCase("14:300.0", 0.013194444)] // Seconds part can be outside of normal range
        [TestCase("140:30.0", 0.097569444)] // Minutes part can be outside of normal range
        [TestCase("30:300.0", 0.024305556)] // Both parts can be outside the range
        [TestCase("140:60.0", null)] // Both hours and minutes are out of range
        [TestCase("60:000.0", null)] // The minutes part starts with 0, but has over 2 digits
        [TestCase("59:300.0", 0.044444444)] // Seconds are added to the minutes, the result is 1:04 minutes
        [TestCase("59:300.59", 0.044451273)] // Can specify 2 digit ms
        [TestCase("00:57.180", 0.000661806)] // Can specify 3 digit ms
        public void TimeSpan_Format47(string timeSpan, double? expectedValue) // 'mm:ss.0'
        {
            AssertCoercion(timeSpan, expectedValue, Tolerance);
        }

        [TestCase("1,000", 1000)]
        [TestCase("1,00", null, Ignore = ".NET parse methods ignores thousands separator, but excel enforces them.")]
        [TestCase("1,000,000", 1000000)]
        [TestCase("1,00,000", null, Ignore = ".NET parse methods ignores thousands separator, but excel enforces them.")]
        [TestCase("(1,000)", -1000)]
        [TestCase("(100)", -100)]
        [TestCase("(-1)", null)]
        public void Number_Format37_38(string number, double? expectedValue) // Format 37+38 '#,##0 ;(#,##0)' '#,##0 ;[Red](#,##0)'
        {
            AssertCoercion(number, expectedValue);
        }

        [TestCase("1,000.15", 1000.15)]
        [TestCase("(1,000.54)", -1000.54)]
        [TestCase("  (   1,000.54  )  ", -1000.54, Ignore = "Excel can parse spaces within braces, but .NET parse method can't.")]
        public void Number_Format39_40(string number, double? expectedValue) // Format 39+40 '#,##0.00;(#,##0.00)'  '#,##0.00;[Red](#,##0.00)'
        {
            AssertCoercion(number, expectedValue);
        }

        [TestCase("1e3", 1000)]
        [TestCase("1e+3", 1000)]
        [TestCase("1e-5", 0.00001)]
        [TestCase("1e0", 1)]
        [TestCase("1.5e2", 150)]
        [TestCase("1e2.5", null)] // Exponent can't be a fraction
        [TestCase("1.52e1", 15.2)]
        [TestCase("-1e2", -100)]
        [TestCase("1E2", 100)]
        public void Number_Format48_11(string number, double? expectedValue) // Format 48+11 '##0.0E+0' '0.00E+00'
        {
            AssertCoercion(number, expectedValue);
        }

        [TestCase("$1", 1)]
        [TestCase("1$", null, Ignore = ".NET parser allows currency symbol at the start or end, but Excel requires correct placement.")]
        [TestCase("($1)", -1)]
        [TestCase("-($1)", null)]
        [TestCase("$100.5", 100.5)]
        [TestCase("$100%", null)]
        [TestCase("($100%)", null)]
        public void Currency(string currency, double? expectedValue) // Currency doesn't have a format in ECMA-376, Part 1, §18.8.30, but VALUE includes currency formats
        {
            AssertCoercion(currency, expectedValue);
        }

        [SetCulture("cs-CZ")]
        [TestCase("$1", null)] // Fallback currency doesn't work nor it should
        [TestCase("Kč 1", null, Ignore = "Excel requires correct placement of currency symbol, while .NET parser accepts any position.")] // incorrect placement
        [TestCase("100.5", null)] // incorrect decimal placement
        [TestCase("10e2 Kč", 1000)]
        [TestCase("30-apr-2000", null)]
        [TestCase("02/28/20", null)]
        [TestCase("10:30 AM", 0.4375)] // AM seems to work for some reason
        [TestCase("10:30 dop.", 0.4375)]
        [TestCase("3-leden", 44564)]
        [TestCase("3-led", 44564)]
        [TestCase("1-leden-2020", 43831)]
        [TestCase("1-led-2020", 43831)]
        [TestCase("led-5", 38353)]
        [TestCase("12:0:18 odp.", 0.50020833333333337d)]
        [TestCase("12:0:18 PM", 0.50020833333333337d)]
        [TestCase("12:0:18 odp", 0.50020833333333337d, Ignore = "Excel can parse even partial PM designator, but .NET requires a full PM designator including the dot at the end.")]
        [TestCase("12:0:18 PM.", 0.50020833333333337d, Ignore = "Excel allows PM designator with a dot at the end.")]
        [TestCase("11/30/2022 25:59", null)]
        [TestCase("25:70,05", 0.018171875)] // For min:sec fraction timespan, both can be over limit, also note use of decimal separator
        public void ParsingTokensAndFormatsDependOnCulture(string currency, double? expectedValue)
        {
            AssertCoercion(currency, expectedValue);
        }

        private static void AssertCoercion(string text, double? expectedValue, double tolerance = 0)
        {
            using var wb = new XLWorkbook();
            var parsedValue = wb.Evaluate($"\"{text}\"*1");
            if (expectedValue is null)
                Assert.AreEqual(XLError.IncompatibleValue, parsedValue);
            else
                Assert.AreEqual(expectedValue.Value, (double)parsedValue, tolerance);
        }
    }
}
