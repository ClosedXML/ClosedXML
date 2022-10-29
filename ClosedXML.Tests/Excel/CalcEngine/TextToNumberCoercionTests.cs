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
        [TestCase("0 1/0", null)]
        public void Fraction_Format12_13(string fraction, double? expectedValue) // Format 12+13 '# ??/??' and  '# ?/?'
        {
            AssertCoercion(fraction, expectedValue);
        }

        [TestCase("00:00", 0)] // Can parse zero
        [TestCase("90:00", 3.75)] // Minutes can be can be over 60
        [TestCase("59:59", 2.499305556)] // Even if looks like mm:ss parsed as h:mm
        [TestCase("10:", 0.416666667)] // Last part can be omitted and zero is used
        [TestCase("9999:", 416.625)] // Upper limit of first part is parseable
        [TestCase("10000:", null)] // Part value over a limit is not parseable
        [TestCase(":5", null)] // Can't omit first part
        [TestCase("24:60", null)] // Only one part can be outside of limit, here are both
        [TestCase("30:59", 1.290972222)] // Hour part over limit
        [TestCase("23:300", 1.166666667)] // Minute part over limit
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
