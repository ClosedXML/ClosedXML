using ClosedXML.Excel;
using NUnit.Framework;
using System;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    [SetCulture("en-US")]
    public class DateAndTimeTests
    {
        [TestCase(2008, 1, 1, ExpectedResult = 39448)]
        [TestCase(2008, 15, 1, ExpectedResult = 39873)]
        [TestCase(2008, -50, 1, ExpectedResult = 37895)]
        [TestCase(2008, 5, 63, ExpectedResult = 39631)]
        [TestCase(2008, 13, 63, ExpectedResult = 39876)]
        [TestCase(2008, 15, -120, ExpectedResult = 39752)]
        [TestCase(1900, 2, 29, ExpectedResult = 60)] // Loveable 29th feb 1900
        [TestCase(1900, 2, 28, ExpectedResult = 59)]
        [TestCase(1900, 1, 1, ExpectedResult = 1)]
        [TestCase(1900, 1, 0, ExpectedResult = 0)] // Excel formats it as 1900-01-00, but more like 1899-12-31
        [TestCase(1899, 1, 1, ExpectedResult = 693598)] // If year < 1900, add 1900 to it
        public double Date_returns_serial_date(int year, int month, int day)
        {
            return XLWorkbook.EvaluateExpr($"DATE({year},{month},{day})").GetNumber();
        }

        [TestCase(1900, 1, -1)] // Serial date -1, below 0
        [TestCase(9999, 12, 32)]
        public void Date_returns_error_when_result_outside_base_date_to_max_date_of_calendar_system(int year, int month, int day)
        {
            var actual = XLWorkbook.EvaluateExpr($"DATE({year},{month},{day})");
            Assert.AreEqual(XLError.NumberInvalid, actual);
        }

        [TestCase(-1, 32000, 1, ExpectedResult = 973586)]  // Year -1.1 behaves as -2
        [TestCase(-1.1, 32000, 1, ExpectedResult = 973221)]
        [TestCase(-2, 32000, 1, ExpectedResult = 973221)]
        [TestCase(2000, -5, 1, ExpectedResult = 36342)] // Month -5.1 behaves as -6
        [TestCase(2000, -5.1, 1, ExpectedResult = 36312)]
        [TestCase(2000, -6, 1, ExpectedResult = 36312)]
        [TestCase(2000, 2, -10, ExpectedResult = 36546)] // Day -10.1 behaves as -11
        [TestCase(2000, 2, -10.1, ExpectedResult = 36545)]
        [TestCase(2000, 2, -11, ExpectedResult = 36545)]
        public double Date_floors_arguments(double year, double month, double day)
        {
            return XLWorkbook.EvaluateExpr($"DATE({year},{month},{day})").GetNumber();
        }

        [TestCase(10000, -32767, 3, "7269-05-03")] // Month can be [-32767..32767)
        [TestCase(10000, -32767.1, 3, XLError.NumberInvalid)]
        [TestCase(2000, 32766.9, 1, "4730-06-01")]
        [TestCase(2000, 32767, 1, XLError.NumberInvalid)]
        [TestCase(2000, 1, 32767.9, "2089-09-16")] // Day is clamped to at most 32767
        [TestCase(2000, 1, 32768, "2089-09-16")]
        [TestCase(2000, 1, 1E+100, "2089-09-16")]
        [TestCase(2000, 1, -32768, "1910-04-14")] // When day is < -32768, day uses 32767 value instead
        [TestCase(2000, 1, -32768.1, "2089-09-16")]
        [TestCase(2000, 1, -1E+100, "2089-09-16")]
        [TestCase(10000, -32000, 1, "7333-04-01")] // Year is clamped to 10000
        [TestCase(10001, -32000, 1, "7333-04-01")]
        [TestCase(1E+100, -32000, 1, "7333-04-01")]
        [TestCase(-1E+100, 1, 1, XLError.NumberInvalid)] // Even if year is less than int.MinValue, there is no error
        public void Date_matches_excel_behavior_for_out_of_range_arguments(double year, double month, double day, object expectedResult)
        {
            if (expectedResult is string iso8601)
                expectedResult = DateTime.Parse(iso8601).ToSerialDateTime();

            var actual = XLWorkbook.EvaluateExpr($"DATE({year},{month},{day})");
            Assert.AreEqual(expectedResult, actual);
        }

        [TestCase("1/1/2006", "12/12/2010", "Y", ExpectedResult = 4)]
        [TestCase("1/1/2006", "12/12/2010", "M", ExpectedResult = 59)]
        [TestCase("1/1/2006", "12/12/2010", "D", ExpectedResult = 1806)]
        [TestCase("1/1/2006", "12/12/2010", "MD", ExpectedResult = 11)]
        [TestCase("1/1/2006", "12/12/2010", "YM", ExpectedResult = 11)]
        [TestCase("1/1/2006", "12/12/2010", "YD", ExpectedResult = 345)]
        [TestCase(38718, 40524, "Y", ExpectedResult = 4)]
        [TestCase(38718, 40524, "M", ExpectedResult = 59)]
        [TestCase(38718, 40524, "D", ExpectedResult = 1806)]
        [TestCase(38718, 40524, "MD", ExpectedResult = 11)]
        [TestCase(38718, 40524, "YM", ExpectedResult = 11)]
        [TestCase(38718, 40524, "YD", ExpectedResult = 345)]
        [TestCase("2001-12-31", "2002-4-15", "YM", ExpectedResult = 3)]
        [TestCase("2001-12-10", "2002-4-15", "YM", ExpectedResult = 4)]
        [TestCase("2001-12-15", "2002-4-15", "YM", ExpectedResult = 4)]
        [TestCase("2001-12-31", "2002-4-15", "YD", ExpectedResult = 105)]
        [TestCase("2001-12-31", "2003-4-15", "YD", ExpectedResult = 105)]
        [TestCase("2002-01-31", "2002-4-15", "YD", ExpectedResult = 74)]
        [TestCase("2001-12-02", "2001-12-15", "Y", ExpectedResult = 0)]
        [TestCase("2001-12-02", "2002-12-02", "Y", ExpectedResult = 1)]
        [TestCase("2006-01-15", "2006-03-14", "M", ExpectedResult = 1)]
        [TestCase("2020-11-22", "2020-11-23 9:00", "D", ExpectedResult = 1)]
        public double Datedif(object startDate, object endDate, string unit)
        {
            if (startDate is string s1) startDate = $"\"{s1}\"";
            if (endDate is string s2) endDate = $"\"{s2}\"";
            return (double)XLWorkbook.EvaluateExpr($"DATEDIF({startDate}, {endDate}, \"{unit}\")");
        }

        [TestCase("\"1/1/2010\"", "\"12/12/2006\"", "Y")]
        [TestCase(40524, 38718, "Y")]
        [TestCase("\"1/1/2006\"", "\"12/12/2010\"", "N")]
        [TestCase(38718, 40524, "N")]
        public void DatedifExceptions(object startDate, object endDate, string unit)
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"DATEDIF({startDate}, {endDate}, \"{unit}\")"));
        }

        [Test]
        public void Datevalue()
        {
            var actual = XLWorkbook.EvaluateExpr("DateValue(\"8/22/2008\")");
            Assert.AreEqual(39682, actual);
        }

        [TestCase(0, ExpectedResult = 0)]
        [TestCase(1, ExpectedResult = 1)]
        [TestCase(31, ExpectedResult = 31)]
        [TestCase(32, ExpectedResult = 1)]
        [TestCase(59, ExpectedResult = 28)]
        [TestCase(60, ExpectedResult = 29)]
        [TestCase(61, ExpectedResult = 1)]
        [TestCase(30000, ExpectedResult = 18)]
        [TestCase(45718, ExpectedResult = 2)]
        public double Day_returns_day_of_a_month_for_serial_culture(double serialDate)
        {
            return XLWorkbook.EvaluateExpr($"DAY({serialDate})").GetNumber();
        }

        [Test]
        public void Day_only_accepts_serial_date_from_0_to_upper_limit_of_calendar_system()
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("DAY(-0.1)"));
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("DAY(DATE(9999,12,31)+1)"));
        }

        [SetCulture("eu-ES")]
        [TestCase("\"2006/1/2 10:45 AM\"", ExpectedResult = 2)]
        [TestCase("DATE(2006,1,2)", ExpectedResult = 2)]
        [TestCase("DATE(2006,0,2)", ExpectedResult = 2)]
        [TestCase("DATE(2013,9,0)", ExpectedResult = 31)]
        public double Day_examples(string date)
        {
            return XLWorkbook.EvaluateExprCurrent($"DAY({date})").GetNumber();
        }

        [Test]
        public void Days()
        {
            var actual = XLWorkbook.EvaluateExpr("DAYS(DATE(2016,10,1),DATE(1992,2,29))");
            Assert.AreEqual(8981, actual);

            actual = XLWorkbook.EvaluateExpr("DAYS(\"2016-10-1\",\"1992-2-29\")");
            Assert.AreEqual(8981, actual);
        }

        [Test]
        public void Days360_uses_US_method_by_default()
        {
            const string formulaFormat = "DAYS360(DATE(2002,2,3),DATE(2005,5,31){0})";
            var defaultResult = XLWorkbook.EvaluateExpr(string.Format(formulaFormat, string.Empty));
            var usResult = XLWorkbook.EvaluateExpr(string.Format(formulaFormat, ",FALSE"));
            var euResult = XLWorkbook.EvaluateExpr(string.Format(formulaFormat, ",TRUE"));
            Assert.AreEqual(1198, defaultResult);
            Assert.AreEqual(usResult, defaultResult);
            Assert.AreNotEqual(euResult, defaultResult);
        }

        [Test]
        public void Days360_Europe1()
        {
            var actual = XLWorkbook.EvaluateExpr("DAYS360(\"1/1/2008\", \"3/31/2008\",TRUE)");
            Assert.AreEqual(89, actual);
        }

        [Test]
        public void Days360_Europe2()
        {
            var actual = XLWorkbook.EvaluateExpr("DAYS360(\"3/31/2008\", \"1/1/2008\",TRUE)");
            Assert.AreEqual(-89, actual);
        }

        [TestCase(2002, 2, 3, 2005, 5, 31, ExpectedResult = 1198)]
        [TestCase(2005, 5, 31, 2002, 2, 3, ExpectedResult = -1197)]
        [TestCase(2008, 1, 1, 2008, 3, 31, ExpectedResult = 90)]
        [TestCase(2008, 3, 31, 2008, 1, 1, ExpectedResult = -89)]
        [TestCase(2020, 2, 29, 2021, 2, 28, ExpectedResult = 358)]
        [TestCase(2020, 5, 29, 2020, 4, 1, ExpectedResult = -58)]
        [TestCase(2020, 5, 29, 2020, 3, 31, ExpectedResult = -58)]
        [TestCase(2020, 5, 30, 2020, 4, 1, ExpectedResult = -59)]
        [TestCase(2020, 5, 30, 2020, 3, 31, ExpectedResult = -60)]
        [TestCase(2020, 5, 30, 2020, 3, 30, ExpectedResult = -60)]
        public double Days360_US_method(int startYear, int startMonth, int startDay, int endYear, int endMonth, int endDay)
        {
            return (double)XLWorkbook.EvaluateExpr($"DAYS360(DATE({startYear},{startMonth},{startDay}),DATE({endYear},{endMonth},{endDay}),FALSE)");
        }

        [TestCase(1900, 2, 27, 1900, 2, 27, ExpectedResult = 0)]
        [TestCase(1900, 2, 27, 1900, 2, 28, ExpectedResult = 1)]
        [TestCase(1900, 2, 27, 1900, 2, 29, ExpectedResult = 2)]
        [TestCase(1900, 2, 27, 1900, 3, 1, ExpectedResult = 4)]
        [TestCase(1900, 2, 28, 1900, 2, 27, ExpectedResult = -1)]
        [TestCase(1900, 2, 28, 1900, 2, 28, ExpectedResult = 0)]
        [TestCase(1900, 2, 28, 1900, 2, 29, ExpectedResult = 1)]
        [TestCase(1900, 2, 28, 1900, 3, 1, ExpectedResult = 3)]
        [TestCase(1900, 2, 29, 1900, 2, 27, ExpectedResult = -3)]
        [TestCase(1900, 2, 29, 1900, 2, 28, ExpectedResult = -2)]
        [TestCase(1900, 2, 29, 1900, 2, 29, ExpectedResult = -1)]
        [TestCase(1900, 2, 29, 1900, 3, 1, ExpectedResult = 1)]
        [TestCase(1900, 3, 1, 1900, 2, 27, ExpectedResult = -4)]
        [TestCase(1900, 3, 1, 1900, 2, 28, ExpectedResult = -3)]
        [TestCase(1900, 3, 1, 1900, 2, 29, ExpectedResult = -2)]
        [TestCase(1900, 3, 1, 1900, 3, 1, ExpectedResult = 0)]
        public double Days360_US_method_for_feb_29_1900(int startYear, int startMonth, int startDay, int endYear, int endMonth, int endDay)
        {
            return (double)XLWorkbook.EvaluateExpr($"DAYS360(DATE({startYear},{startMonth},{startDay}),DATE({endYear},{endMonth},{endDay}),FALSE)");
        }

        [Test]
        public void EDate_Negative1()
        {
            var actual = XLWorkbook.EvaluateExpr("EDate(\"3/1/2008\", -1)");
            Assert.AreEqual(new DateTime(2008, 2, 1).ToSerialDateTime(), actual);
        }

        [Test]
        public void EDate_Negative2()
        {
            var actual = XLWorkbook.EvaluateExpr("EDate(\"3/31/2008\", -1)");
            Assert.AreEqual(new DateTime(2008, 2, 29).ToSerialDateTime(), actual);
        }

        [Test]
        public void EDate_Positive1()
        {
            var actual = XLWorkbook.EvaluateExpr("EDate(\"3/1/2008\", 1)");
            Assert.AreEqual(new DateTime(2008, 4, 1).ToSerialDateTime(), actual);
        }

        [Test]
        public void EDate_Positive2()
        {
            var actual = XLWorkbook.EvaluateExpr("EDate(\"3/31/2008\", 1)");
            Assert.AreEqual(new DateTime(2008, 4, 30).ToSerialDateTime(), actual);
        }

        [Test]
        public void EOMonth_Negative()
        {
            var actual = XLWorkbook.EvaluateExpr("EOMonth(\"3/1/2008\", -1)");
            Assert.AreEqual(new DateTime(2008, 2, 29).ToSerialDateTime(), actual);
        }

        [Test]
        public void EOMonth_Positive()
        {
            var actual = XLWorkbook.EvaluateExpr("EOMonth(\"3/31/2008\", 1)");
            Assert.AreEqual(new DateTime(2008, 4, 30).ToSerialDateTime(), actual);
        }

        [TestCase("0", ExpectedResult = 0)]
        [TestCase("0.25", ExpectedResult = 6)]
        [TestCase("0.5", ExpectedResult = 12)]
        [TestCase("0.75", ExpectedResult = 18)]
        [TestCase("1", ExpectedResult = 0)]
        [TestCase("1.75", ExpectedResult = 18)]
        [TestCase("\"7/18/2011 7:45\"", ExpectedResult = 7)]
        [TestCase("\"4/21/2012\"", ExpectedResult = 0)]
        [TestCase("\"12:00:00\"", ExpectedResult = 12)]
        [TestCase("\"8/22/2008 3:30:45 PM\"", ExpectedResult = 15, Ignore = "We don't parse seconds")]
        [TestCase("\"8/22/2008 3:30 PM\"", ExpectedResult = 15)]
        [TestCase("DATE(2006,2,26)+TIME(2,10,20)", ExpectedResult = 2)]
        [TestCase("TIME(22,56,34)", ExpectedResult = 22)]
        [TestCase("\"22-Oct-2001 10:53:12\"", ExpectedResult = 10, Ignore = "We don't parse seconds plus culture is wrong")]
        [TestCase("\"October 22, 2001 10:53\"", ExpectedResult = 10)]
        [TestCase("\"10:53:12 pm\"", ExpectedResult = 22)]
        [TestCase("\"22:53:12\"", ExpectedResult = 22)]
        public double Hour_returns_hour_of_serial_date(string dateArg)
        {
            return XLWorkbook.EvaluateExprCurrent($"HOUR({dateArg})").GetNumber();
        }

        [Test]
        public void Hour_accepts_only_serial_time_between_zero_and_upper_limit_of_date_system()
        {
            Assert.AreEqual(0, XLWorkbook.EvaluateExprCurrent("HOUR(0)"));
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExprCurrent("HOUR(-0.1)"));

            Assert.AreEqual(21, XLWorkbook.EvaluateExprCurrent("HOUR(DATE(9999,12,31)+0.9)"));
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExprCurrent("HOUR(DATE(9999,12,31)+1)"));
        }

        [TestCase("0", ExpectedResult = 0)]
        [TestCase("0.5", ExpectedResult = 0)]
        [TestCase("0.68", ExpectedResult = 19)]
        [TestCase("0.69", ExpectedResult = 33)]
        [TestCase("0.85", ExpectedResult = 24)]
        [TestCase("10.85", ExpectedResult = 24)]
        [TestCase("\"14:47:20\"", ExpectedResult = 47)]
        [TestCase("\"8/22/2008 3:30 AM\"", ExpectedResult = 30)]
        public double Minute_returns_minute_of_serial_date(string dateArg)
        {
            return XLWorkbook.EvaluateExprCurrent($"MINUTE({dateArg})").GetNumber();
        }

        [Test]
        public void Minute_accepts_only_serial_time_between_zero_and_upper_limit_of_date_system()
        {
            Assert.AreEqual(0, XLWorkbook.EvaluateExprCurrent("MINUTE(0)"));
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExprCurrent("MINUTE(-0.1)"));

            Assert.AreEqual(36, XLWorkbook.EvaluateExprCurrent("MINUTE(DATE(9999,12,31)+0.9)"));
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExprCurrent("MINUTE(DATE(9999,12,31)+1)"));
        }

        [SetCulture("eu-ES")]
        [TestCase(0, ExpectedResult = 1)] // 1900-01-00
        [TestCase(31, ExpectedResult = 1)] // 1900-01-31
        [TestCase(32, ExpectedResult = 2)] // 1900-02-01
        [TestCase(59, ExpectedResult = 2)] // 1900-02-28
        [TestCase(60, ExpectedResult = 2)] // 1900-02-29
        [TestCase(61, ExpectedResult = 3)] // 1900-03-01
        [TestCase("DATE(2006,1,2)", ExpectedResult = 1)]
        [TestCase("DATE(2006,0,2)", ExpectedResult = 12)]
        [TestCase("\"2006/1/2 10:45 AM\"", ExpectedResult = 1)]
        [TestCase("30000", ExpectedResult = 2)]
        [TestCase("45596", ExpectedResult = 10)]
        [TestCase("45596.9", ExpectedResult = 10)]
        [TestCase("45597", ExpectedResult = 11)]
        public double Month_returns_month_of_serial_date(object argument)
        {
            return XLWorkbook.EvaluateExprCurrent($"MONTH({argument})").GetNumber();
        }

        [Test]
        public void Month_serial_date_must_be_between_zero_and_upper_limit_of_date_system()
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("MONTH(-0.1)"));
            Assert.AreEqual(12, XLWorkbook.EvaluateExpr("MONTH(DATE(9999,12,31) + 0.9)"));
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("MONTH(DATE(9999,12,31) + 1)"));
        }

        [Test]
        public void IsoWeekNum()
        {
            var actual = XLWorkbook.EvaluateExpr("ISOWEEKNUM(DATEVALUE(\"2012-3-9\"))");
            Assert.AreEqual(10, actual);

            actual = XLWorkbook.EvaluateExpr("ISOWEEKNUM(DATE(2012,12,31))");
            Assert.AreEqual(1, actual);
        }

        [Test]
        public void NetWorkDays_with_holidays()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.FirstCell().SetValue("Date")
                .CellBelow().SetValue(new DateTime(2008, 10, 1))
                .CellBelow().SetValue(new DateTime(2009, 3, 1))
                .CellBelow().SetValue(new DateTime(2008, 11, 26))
                .CellBelow().SetValue(new DateTime(2008, 12, 4))
                .CellBelow().SetValue(new DateTime(2009, 1, 21))
                .CellBelow().SetValue(new DateTime(2009, 1, 4)) // Holiday is on Sunday - do not count twice
                .CellBelow().SetValue(new DateTime(2009, 1, 6))  // Workweek holiday is specified twice, shouldn't be counted twice
                .CellBelow().SetValue(new DateTime(2009, 1, 6))
                .CellBelow().SetValue(new DateTime(2008, 9, 30)) // Tuesday holiday just before the first date, shouldn't be counted
                .CellBelow().SetValue(new DateTime(2009, 3, 2)) // Monday holiday just after the last date, shouldn't be counted
                ;
            var actual = ws.Evaluate("NETWORKDAYS(A2, A3, A4:A11)");
            Assert.AreEqual(104, actual);
        }

        [TestCase("2024-10-01", "2024-10-01", 1)] // Tue-Tue
        [TestCase("2024-10-01", "2024-10-02", 2)] // Tue-Wed
        [TestCase("2024-10-01", "2024-10-03", 3)] // Tue-Thu
        [TestCase("2024-10-01", "2024-10-04", 4)] // Tue-Fri
        [TestCase("2024-10-01", "2024-10-05", 4)] // Tue-Sat
        [TestCase("2024-10-01", "2024-10-06", 4)] // Tue-Sun
        [TestCase("2024-10-01", "2024-10-07", 5)] // Tue-Mon
        [TestCase("2024-09-29", "2024-10-12", 10)] // Sun-Sat
        [TestCase("2024-09-29", "2024-10-13", 10)] // Sun-Sun
        [TestCase("2024-09-29", "2024-10-14", 11)] // Sun-Mon
        [TestCase("2024-09-29", "2024-10-15", 12)] // Sun-Tue
        [TestCase("2024-09-29", "2024-10-16", 13)] // Sun-Wed
        [TestCase("2024-09-29", "2024-10-17", 14)] // Sun-Thu
        [TestCase("2024-09-29", "2024-10-18", 15)] // Sun-Fri
        [TestCase("2024-09-29", "2024-10-19", 15)] // Sun-Sat
        public void NetWorkDays_non_full_weeks_are_counted_correctly(string startDate, string endDate, int expected)
        {
            var actual = XLWorkbook.EvaluateExpr($"NETWORKDAYS(\"{startDate}\", \"{endDate}\")");
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Culture("en-US")]
        public void NetWorkDays_with_end_date_earlier_than_start_date()
        {
            var actual = XLWorkbook.EvaluateExpr("NETWORKDAYS(\"3/01/2009\", \"10/01/2008\")");
            Assert.AreEqual(-108, actual);

            actual = XLWorkbook.EvaluateExpr("NETWORKDAYS(\"2016-01-01\", \"2015-12-23\")");
            Assert.AreEqual(-8, actual);
        }

        [Test]
        [Culture("en-US")]
        public void NetWorkDays_behavior()
        {
            using var wb = new XLWorkbook();
            var actual = wb.Evaluate("NETWORKDAYS(\"10/01/2008\", \"3/01/2009\", \"11/26/2008\")");
            Assert.AreEqual(107, actual);

            // Example from specification. Except spec wrong. The value is 1 off from Excel value.
            Assert.AreEqual(22, wb.Evaluate("NETWORKDAYS(DATE(2006, 1, 1), DATE(2006, 1, 31))"));
            Assert.AreEqual(-22, wb.Evaluate("NETWORKDAYS(DATE(2006, 1, 31), DATE(2006, 1, 1))"));
            Assert.AreEqual(21, wb.Evaluate("NETWORKDAYS(DATE(2006, 1, 1), DATE(2006, 2, 1), { \"2006-01-02\", \"2006-01-16\" })"));

            // Scalar number is accepted for holidays
            Assert.AreEqual(6, wb.Evaluate("NETWORKDAYS(1, 10, 2)"));

            // Scalar logical causes conversion error
            Assert.AreEqual(XLError.IncompatibleValue, wb.Evaluate("NETWORKDAYS(TRUE, 10)"));
            Assert.AreEqual(XLError.IncompatibleValue, wb.Evaluate("NETWORKDAYS(0, TRUE)"));
            Assert.AreEqual(XLError.IncompatibleValue, wb.Evaluate("NETWORKDAYS(1, 10, TRUE)"));

            // Scalar text is converted
            Assert.AreEqual(6, wb.Evaluate("NETWORKDAYS(\"1\", \"10\", \"2\")"));
            Assert.AreEqual(6, wb.Evaluate("NETWORKDAYS(1, 10, \"0 4/2\")"));
            Assert.AreEqual(6, wb.Evaluate("NETWORKDAYS(1, 10, \"1900-01-02\")"));
            Assert.AreEqual(XLError.IncompatibleValue, wb.Evaluate("NETWORKDAYS(\"Text\", 10)"));
            Assert.AreEqual(XLError.IncompatibleValue, wb.Evaluate("NETWORKDAYS(1, \"Text\")"));
            Assert.AreEqual(XLError.IncompatibleValue, wb.Evaluate("NETWORKDAYS(1, 10, \"Text\")"));

            // Array accepts numbers and converts text
            Assert.AreEqual(5, wb.Evaluate("NETWORKDAYS(1, 10, {\"2\", 3})"));
            Assert.AreEqual(XLError.IncompatibleValue, wb.Evaluate("NETWORKDAYS(1, 10, {\"Text\"})"));
            Assert.AreEqual(XLError.IncompatibleValue, wb.Evaluate("NETWORKDAYS(1, 10, {TRUE})"));

            // Same conversion logic applies to reference values
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = Blank.Value; // Ignored
            ws.Cell("A2").Value = false; // Causes conversion error
            ws.Cell("A3").Value = true; // Causes conversion error
            ws.Cell("A4").Value = 37147; // 2001-09-13
            ws.Cell("A5").Value = "2001-09-12"; // Monday
            ws.Cell("A6").Value = XLError.NoValueAvailable;

            Assert.AreEqual(175, ws.Evaluate("NETWORKDAYS(\"2001-05-01\", \"2001-12-31\", A1)"));
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("NETWORKDAYS(\"2001-05-01\", \"2001-12-31\", A1:A3)"));
            Assert.AreEqual(173, ws.Evaluate("NETWORKDAYS(\"2001-05-01\",\"2001-12-31\", A4:A5)"));

            // Errors are propagated
            Assert.AreEqual(XLError.NoValueAvailable, wb.Evaluate("NETWORKDAYS(#N/A, 10)"));
            Assert.AreEqual(XLError.NoValueAvailable, wb.Evaluate("NETWORKDAYS(1, #N/A)"));
            Assert.AreEqual(XLError.NoValueAvailable, wb.Evaluate("NETWORKDAYS(1, 10, {#N/A})"));
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate("NETWORKDAYS(1, 10, A6)"));
        }

        [TestCase("0", ExpectedResult = 0)]
        [TestCase("\"3:30:45\"", ExpectedResult = 45)]
        public double Second_returns_minute_of_serial_date(string dateArg)
        {
            return XLWorkbook.EvaluateExprCurrent($"SECOND({dateArg})").GetNumber();
        }

        [Test]
        public void Second_accepts_only_serial_time_between_zero_and_upper_limit_of_date_system()
        {
            Assert.AreEqual(0, XLWorkbook.EvaluateExprCurrent("SECOND(0)"));
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExprCurrent("SECOND(-0.1)"));

            Assert.AreEqual(51, XLWorkbook.EvaluateExprCurrent("SECOND(DATE(9999,12,31)+0.9999)"));
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExprCurrent("SECOND(DATE(9999,12,31)+1)"));
        }

        [TestCase(0, 0, 0, ExpectedResult = 0)]
        [TestCase(0, 0, 1, ExpectedResult = 0.0000115740740741)]
        [TestCase(0, 0, 2, ExpectedResult = 0.0000231481481481)]
        [TestCase(0, 0, 20, ExpectedResult = 0.0002314814814815)]
        [TestCase(2, 3, 20, ExpectedResult = 0.0856481481481481)]
        [TestCase(12, 0, 0, ExpectedResult = 0.5000000000000000)]
        [TestCase(23, 59, 59, ExpectedResult = 0.9999884259259260)]
        [TestCase(26, 120, 240, ExpectedResult = 0.1694444444444450)]
        [TestCase(1, 2, 3, ExpectedResult = 0.043090277777778)]
        [TestCase(1.9, 2.9, 3.9, ExpectedResult = 0.043090277777778)]
        [TestCase(24, 0, 0, ExpectedResult = 0)]
        [TestCase(0, 42 * 60, 0, ExpectedResult = 0.75)]
        [TestCase(0, 0, 60 * 60 * 3, ExpectedResult = 0.125)]
        [TestCase(120, 240, 347, ExpectedResult = 0.170682870370)]
        [DefaultFloatingPointTolerance(XLHelper.Epsilon)]
        public double Time_returns_serial_date_time(double hour, double minute, double second)
        {
            return (double)XLWorkbook.EvaluateExpr($"TIME({hour},{minute},{second})");
        }

        [TestCase(-0.1, 0, 0)]
        [TestCase(32768, 0, 0)]
        [TestCase(0, -0.1, 0)]
        [TestCase(0, 32768, 0)]
        [TestCase(0, 0, -0.1)]
        [TestCase(0, 0, 32768)]
        public void Time_components_must_be_in_zero_to_32767_interval(double hour, double minute, double second)
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"TIME({hour},{minute},{second})"));
        }

        [Test]
        public void TimeValue1()
        {
            var actual = (double)XLWorkbook.EvaluateExpr("TimeValue(\"2:24 AM\")");
            Assert.IsTrue(XLHelper.AreEqual(0.1, actual));
        }

        [Test]
        public void TimeValue2()
        {
            var actual = (double)XLWorkbook.EvaluateExpr("TimeValue(\"22-Aug-2008 6:35 AM\")");
            Assert.IsTrue(XLHelper.AreEqual(0.27430555555555558, actual));
        }

        [Test]
        public void Today()
        {
            var actual = (double)XLWorkbook.EvaluateExpr("TODAY()");
            Assert.AreEqual(DateTime.Today.ToSerialDateTime(), actual);
        }

        [TestCase("\"2/14/2008\"", 1, 5)]
        [TestCase("\"2/14/2008\"", 2, 4)]
        [TestCase("\"2/14/2008\"", 3, 3)]
        [TestCase("\"2/14/2008\"", 11, 4)]
        [TestCase("\"2/14/2008\"", 12, 3)]
        [TestCase("\"2/14/2008\"", 13, 2)]
        [TestCase("\"2/14/2008\"", 14, 1)]
        [TestCase("\"2/14/2008\"", 15, 7)]
        [TestCase("\"2/14/2008\"", 16, 6)]
        [TestCase("\"2/14/2008\"", 17, 5)]
        public void Weekday_calculates_week_day(string value, int flag, int expected)
        {
            var actual = XLWorkbook.EvaluateExpr($"WEEKDAY({value}, {flag})");
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void Weekday_without_flag()
        {
            var actual = XLWorkbook.EvaluateExpr("WEEKDAY(\"2/14/2008\")");
            Assert.AreEqual(5, actual);
        }

        [Test]
        public void Weekday_behavior()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            ws.Cell("A1").Value = 45577;
            Assert.AreEqual(7, ws.Evaluate("WEEKDAY(A1)"));

            // Time of the day doesn't matter, serial date is truncated
            Assert.AreEqual(7, XLWorkbook.EvaluateExpr("WEEKDAY(45577.9, 1.9)"));

            Assert.AreEqual(7, XLWorkbook.EvaluateExpr("WEEKDAY(0)"));
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("WEEKDAY(-1)"));

            // Year 10k
            Assert.AreEqual(6, XLWorkbook.EvaluateExpr("WEEKDAY(2958465)"));
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("WEEKDAY(2958466)"));

            // Convert from logical/text to number
            Assert.AreEqual(1, XLWorkbook.EvaluateExpr("WEEKDAY(TRUE)"));
            Assert.AreEqual(1, XLWorkbook.EvaluateExpr("WEEKDAY(\"0 2/2\")"));
            Assert.AreEqual(1, XLWorkbook.EvaluateExpr("WEEKDAY(1, TRUE)"));
            Assert.AreEqual(1, XLWorkbook.EvaluateExpr("WEEKDAY(1, \"1 0/2\")"));
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("WEEKDAY(\"text\")"));
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("WEEKDAY(1, \"text\")"));

            // Flag can only have some values
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("WEEKDAY(1, 0)"));
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("WEEKDAY(1, 4)"));
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("WEEKDAY(1, 10)"));
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("WEEKDAY(1, 18)"));

            // Error is propagated
            Assert.AreEqual(XLError.NoValueAvailable, XLWorkbook.EvaluateExpr("WEEKDAY(#N/A)"));
            Assert.AreEqual(XLError.NoValueAvailable, XLWorkbook.EvaluateExpr("WEEKDAY(5, #N/A)"));
        }

        [Test]
        public void Weeknum_1()
        {
            Assert.AreEqual(11, XLWorkbook.EvaluateExpr("Weeknum(\"3/9/2000\", 1)"));
        }

        [Test]
        public void Weeknum_10()
        {
            Assert.AreEqual(11, XLWorkbook.EvaluateExpr("Weeknum(\"3/9/2004\", 2)"));
        }

        [Test]
        public void Weeknum_11()
        {
            Assert.AreEqual(11, XLWorkbook.EvaluateExpr("Weeknum(\"3/9/2005\", 1)"));
        }

        [Test]
        public void Weeknum_12()
        {
            Assert.AreEqual(11, XLWorkbook.EvaluateExpr("Weeknum(\"3/9/2005\", 2)"));
        }

        [Test]
        public void Weeknum_13()
        {
            Assert.AreEqual(10, XLWorkbook.EvaluateExpr("Weeknum(\"3/9/2006\", 1)"));
        }

        [Test]
        public void Weeknum_14()
        {
            Assert.AreEqual(11, XLWorkbook.EvaluateExpr("Weeknum(\"3/9/2006\", 2)"));
        }

        [Test]
        public void Weeknum_15()
        {
            Assert.AreEqual(10, XLWorkbook.EvaluateExpr("Weeknum(\"3/9/2007\", 1)"));
        }

        [Test]
        public void Weeknum_16()
        {
            Assert.AreEqual(10, XLWorkbook.EvaluateExpr("Weeknum(\"3/9/2007\", 2)"));
        }

        [Test]
        public void Weeknum_17()
        {
            Assert.AreEqual(11, XLWorkbook.EvaluateExpr("Weeknum(\"3/9/2008\", 1)"));
        }

        [Test]
        public void Weeknum_18()
        {
            Assert.AreEqual(10, XLWorkbook.EvaluateExpr("Weeknum(\"3/9/2008\", 2)"));
        }

        [Test]
        public void Weeknum_19()
        {
            Assert.AreEqual(11, XLWorkbook.EvaluateExpr("Weeknum(\"3/9/2009\", 1)"));
        }

        [Test]
        public void Weeknum_2()
        {
            Assert.AreEqual(11, XLWorkbook.EvaluateExpr("Weeknum(\"3/9/2000\", 2)"));
        }

        [Test]
        public void Weeknum_20()
        {
            Assert.AreEqual(11, XLWorkbook.EvaluateExpr("Weeknum(\"3/9/2009\", 2)"));
        }

        [Test]
        public void Weeknum_3()
        {
            Assert.AreEqual(10, XLWorkbook.EvaluateExpr("Weeknum(\"3/9/2001\", 1)"));
        }

        [Test]
        public void Weeknum_4()
        {
            Assert.AreEqual(10, XLWorkbook.EvaluateExpr("Weeknum(\"3/9/2001\", 2)"));
        }

        [Test]
        public void Weeknum_5()
        {
            Assert.AreEqual(10, XLWorkbook.EvaluateExpr("Weeknum(\"3/9/2002\", 1)"));
        }

        [Test]
        public void Weeknum_6()
        {
            Assert.AreEqual(10, XLWorkbook.EvaluateExpr("Weeknum(\"3/9/2002\", 2)"));
        }

        [Test]
        public void Weeknum_7()
        {
            Assert.AreEqual(11, XLWorkbook.EvaluateExpr("Weeknum(\"3/9/2003\", 1)"));
        }

        [Test]
        public void Weeknum_8()
        {
            Assert.AreEqual(10, XLWorkbook.EvaluateExpr("Weeknum(\"3/9/2003\", 2)"));
        }

        [Test]
        public void Weeknum_9()
        {
            Assert.AreEqual(11, XLWorkbook.EvaluateExpr("Weeknum(\"3/9/2004\", 1)"));
        }

        [Test]
        public void Weeknum_Default()
        {
            var actual = XLWorkbook.EvaluateExpr("Weeknum(\"3/9/2008\")");
            Assert.AreEqual(11, actual);
        }

        [Test]
        public void Workdays_MultipleHolidaysGiven()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Date")
                .CellBelow().SetValue(new DateTime(2008, 10, 1))
                .CellBelow().SetValue(151)
                .CellBelow().SetValue(new DateTime(2008, 11, 26))
                .CellBelow().SetValue(new DateTime(2008, 12, 4))
                .CellBelow().SetValue(new DateTime(2009, 1, 21));
            var actual = ws.Evaluate("Workday(A2,A3,A4:A6)");
            Assert.AreEqual(new DateTime(2009, 5, 5).ToSerialDateTime(), actual);
        }

        [Test]
        public void Workdays_NoHolidaysGiven()
        {
            var actual = XLWorkbook.EvaluateExpr("Workday(\"10/01/2008\", 151)");
            Assert.AreEqual(new DateTime(2009, 4, 30).ToSerialDateTime(), actual);

            actual = XLWorkbook.EvaluateExpr("Workday(\"2016-01-01\", -10)");
            Assert.AreEqual(new DateTime(2015, 12, 18).ToSerialDateTime(), actual);
        }

        [Test]
        public void Workdays_OneHolidaysGiven()
        {
            var actual = XLWorkbook.EvaluateExpr("Workday(\"10/01/2008\", 152, \"11/26/2008\")");
            Assert.AreEqual(new DateTime(2009, 5, 4).ToSerialDateTime(), actual);
        }

        [TestCase(0, 0, 0)]
        [TestCase(0, 1, 2)]
        [TestCase(1, 1, 2)]
        [TestCase(2, 1, 3)]
        [TestCase(0, 5, 6)]
        [TestCase(2, 8, 12)]
        [TestCase(10, -1, 9)]
        [TestCase(10, -2, 6)]
        [TestCase(10, -3, 5)]
        [TestCase(9, -1, 6)]
        [TestCase(9, -2, 5)]
        [TestCase(8, -1, 6)]
        [TestCase(7, -1, 6)]
        [TestCase(6, -1, 5)]
        public void Workdays(int startDate, int dayOffset, int expected)
        {
            var actual = XLWorkbook.EvaluateExpr($"WORKDAY({startDate}, {dayOffset})");
            Assert.AreEqual(expected, actual);
        }

        [TestCase(0, 1, new[] { 1 }, 2)]
        [TestCase(0, 1, new[] { 2 }, 3)]
        [TestCase(0, 5, new[] { 2, 4 }, 10)]
        [TestCase(0, 4, new[] { 2, 4, 6 }, 10)]
        [TestCase(0, 3, new[] { 2, 3, 4, 6 }, 10)]
        [TestCase(0, 2, new[] { 2, 3, 4, 5, 6 }, 10)]
        [TestCase(0, 1, new[] { 2, 3, 5 }, 4)]
        [TestCase(0, 2, new[] { 2, 3, 5 }, 6)]
        [TestCase(2, 1, new[] { 2 }, 3)]
        [TestCase(15, -1, new[] { 13 }, 12)] // 15 = Sunday
        [TestCase(100, -5, new[] { 82, 93, 94, 95, 94, 100 }, 88)]
        [TestCase(98, -2, new[] { 97 }, 95)]
        public void Workdays_with_holiday(int startDate, int dayOffset, int[] holidays, int expected)
        {
            var actual = XLWorkbook.EvaluateExpr($"WORKDAY({startDate}, {dayOffset}, {{{string.Join(",", holidays)}}})");
            Assert.AreEqual(expected, actual);
        }

        [TestCase("\"8/22/2008\"", 2008)]
        [TestCase("\"1/2/2006 10:45 AM\"", 2006)]
        [TestCase("0", 1900)]
        [TestCase("0.5", 1900)]
        [TestCase("1", 1900)]
        [TestCase("59", 1900)]
        [TestCase("60", 1900)]
        [TestCase("366", 1900)]
        [TestCase("367", 1901)]
        [TestCase("DATE(9999,12,31)+0.9", 9999)]
        [TestCase("DATE(9999,12,31)+1", XLError.NumberInvalid)]
        [TestCase("-1", XLError.NumberInvalid)]
        [TestCase("\"test\"", XLError.IncompatibleValue)]
        [TestCase("IF(TRUE,)", 1900)] // Blank
        [TestCase("TRUE", 1900)]
        [TestCase("FALSE", 1900)]
        [TestCase("#DIV/0!", XLError.DivisionByZero)]
        public void Year(string value, object expected)
        {
            var actual = XLWorkbook.EvaluateExpr($"YEAR({value})");
            Assert.AreEqual(XLCellValue.FromObject(expected), actual);
        }

        [Test]
        public void Year_BlankValue()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = Blank.Value;
            ws.Cell("A2").FormulaA1 = "YEAR(A1)";
            var valueA2 = ws.Cell("A2").Value;
            Assert.AreEqual(1900, valueA2);
        }

        [DefaultFloatingPointTolerance(XLHelper.Epsilon)]
        [TestCase(0, 2008, 1, 1, 2008, 3, 31, ExpectedResult = 0.25)]
        [TestCase(0, 2008, 1, 1, 2013, 3, 31, ExpectedResult = 5.25)]
        [TestCase(1, 2008, 1, 1, 2008, 3, 31, ExpectedResult = 0.24590163934426229)]
        [TestCase(1, 2008, 1, 1, 2013, 3, 31, ExpectedResult = 5.24452554744526)]
        [TestCase(1, 1900, 1, 10, 2024, 2, 29, ExpectedResult = 124.137572279657)]
        [TestCase(1, 1924, 6, 25, 2025, 2, 28, ExpectedResult = 100.67763581705)]
        [TestCase(2, 2008, 1, 1, 2008, 3, 31, ExpectedResult = 0.25)]
        [TestCase(2, 2008, 1, 1, 2013, 3, 31, ExpectedResult = 5.32222222222222)]
        [TestCase(3, 2008, 1, 1, 2008, 3, 31, ExpectedResult = 0.24657534246575341)]
        [TestCase(3, 2008, 1, 1, 2013, 3, 31, ExpectedResult = 5.24931506849315)]
        [TestCase(4, 2008, 1, 1, 2008, 3, 31, ExpectedResult = 0.24722222222222223)]
        [TestCase(4, 2008, 1, 1, 2013, 3, 31, ExpectedResult = 5.24722222222222)]
        [TestCase(0, 2006, 1, 1, 2006, 3, 26, ExpectedResult = 0.23611111111)]
        [TestCase(0, 2006, 3, 26, 2006, 1, 1, ExpectedResult = 0.23611111111)]
        [TestCase(0, 2006, 1, 1, 2006, 7, 1, ExpectedResult = 0.5)]
        [TestCase(0, 2006, 1, 1, 2007, 9, 1, ExpectedResult = 1.6666666666)]
        [TestCase(1, 2006, 1, 1, 2006, 7, 1, ExpectedResult = 0.495890411)]
        [TestCase(2, 2006, 1, 1, 2006, 7, 1, ExpectedResult = 0.5027777778)]
        [TestCase(3, 2006, 1, 1, 2006, 7, 1, ExpectedResult = 0.495890411)]
        [TestCase(4, 2006, 1, 1, 2006, 7, 1, ExpectedResult = 0.5)]
        [TestCase(1, 2004, 3, 1, 2006, 3, 1, ExpectedResult = 1.9981751825)]
        public double Yearfrac_calculates_fraction_of_a_year(double basis, double startYear, double startMonth, double startDay, double endYear, double endMonth, double endDay)
        {
            return (double)XLWorkbook.EvaluateExpr($"YEARFRAC(DATE({startYear},{startMonth},{startDay}),DATE({endYear},{endMonth},{endDay}),{basis})");
        }
    }
}
