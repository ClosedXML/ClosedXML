using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    public class TextToNumberCoercionTests
    {
        [Test]
        public void TimeSpan_MaximumResolutionIsOneMs()
        {
            var firstValue = (double)XLWorkbook.EvaluateExpr("\"0:0:0.0015\" * 1");
            var secondValue = (double)XLWorkbook.EvaluateExpr("\"0:0:0.0024\" * 1");
            Assert.AreEqual(firstValue, secondValue);
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
            var parsedValue = XLWorkbook.EvaluateExpr($"\"{timeSpan}\"*1");
            if (expectedValue is null)
                Assert.AreEqual(XLError.IncompatibleValue, parsedValue);
            else
                Assert.AreEqual(expectedValue.Value, (double)parsedValue, 0.000001);
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
            var parsedValue = XLWorkbook.EvaluateExpr($"\"{timeSpan}\"*1");
            if (expectedValue is null)
                Assert.AreEqual(XLError.IncompatibleValue, parsedValue);
            else
                Assert.AreEqual(expectedValue.Value, (double)parsedValue, 0.000001);
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
            var parsedValue = XLWorkbook.EvaluateExpr($"\"{timeSpan}\"*1");
            if (expectedValue is null)
                Assert.AreEqual(XLError.IncompatibleValue, parsedValue);
            else
                Assert.AreEqual(expectedValue.Value, (double)parsedValue, 0.000001);
        }

    }
}
