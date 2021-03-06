using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Globalization;
using System.Threading;

namespace ClosedXML.Tests.Excel.DataValidations
{
    [TestFixture]
    public class DateAndTimeTests
    {
        [SetUp]
        public void SetCultureInfo()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");
        }

        [Test]
        public void Date()
        {
            Object actual;

            actual = XLWorkbook.EvaluateExpr("Date(2008, 1, 1)");
            Assert.AreEqual(39448, actual);

            actual = XLWorkbook.EvaluateExpr("Date(2008, 15, 1)");
            Assert.AreEqual(39873, actual);

            actual = XLWorkbook.EvaluateExpr("Date(2008, -50, 1)");
            Assert.AreEqual(37895, actual);

            actual = XLWorkbook.EvaluateExpr("Date(2008, 5, 63)");
            Assert.AreEqual(39631, actual);

            actual = XLWorkbook.EvaluateExpr("Date(2008, 13, 63)");
            Assert.AreEqual(39876, actual);

            actual = XLWorkbook.EvaluateExpr("Date(2008, 15, -120)");
            Assert.AreEqual(39752, actual);
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
            Assert.AreEqual(XLCalculationErrorType.NumberInvalid, XLWorkbook.EvaluateExpr($"DATEDIF({startDate}, {endDate}, \"{unit}\")"));
        }

        [Test]
        public void Datevalue()
        {
            Object actual = XLWorkbook.EvaluateExpr("DateValue(\"8/22/2008\")");
            Assert.AreEqual(39682, actual);
        }

        [Test]
        public void Day()
        {
            Object actual = XLWorkbook.EvaluateExpr("Day(\"8/22/2008\")");
            Assert.AreEqual(22, actual);
        }

        [Test]
        public void Days()
        {
            Object actual = XLWorkbook.EvaluateExpr("DAYS(DATE(2016,10,1),DATE(1992,2,29))");
            Assert.AreEqual(8981, actual);

            actual = XLWorkbook.EvaluateExpr("DAYS(\"2016-10-1\",\"1992-2-29\")");
            Assert.AreEqual(8981, actual);
        }

        [Test]
        public void DayWithDifferentCulture()
        {
            CultureInfo ci = new CultureInfo(CultureInfo.InvariantCulture.LCID);
            ci.DateTimeFormat.ShortDatePattern = "dd/MM/yyyy";
            Thread.CurrentThread.CurrentCulture = ci;
            Object actual = XLWorkbook.EvaluateExpr("Day(\"1/6/2008\")");
            Assert.AreEqual(1, actual);
        }

        [Test]
        public void Days360_Default()
        {
            Object actual = XLWorkbook.EvaluateExpr("Days360(\"1/30/2008\", \"2/1/2008\")");
            Assert.AreEqual(1, actual);
        }

        [Test]
        public void Days360_Europe1()
        {
            Object actual = XLWorkbook.EvaluateExpr("DAYS360(\"1/1/2008\", \"3/31/2008\",TRUE)");
            Assert.AreEqual(89, actual);
        }

        [Test]
        public void Days360_Europe2()
        {
            Object actual = XLWorkbook.EvaluateExpr("DAYS360(\"3/31/2008\", \"1/1/2008\",TRUE)");
            Assert.AreEqual(-89, actual);
        }

        [Test]
        public void Days360_US1()
        {
            Object actual = XLWorkbook.EvaluateExpr("DAYS360(\"1/1/2008\", \"3/31/2008\",FALSE)");
            Assert.AreEqual(90, actual);
        }

        [Test]
        public void Days360_US2()
        {
            Object actual = XLWorkbook.EvaluateExpr("DAYS360(\"3/31/2008\", \"1/1/2008\",FALSE)");
            Assert.AreEqual(-89, actual);
        }

        [Test]
        public void EDate_Negative1()
        {
            Object actual = XLWorkbook.EvaluateExpr("EDate(\"3/1/2008\", -1)");
            Assert.AreEqual(new DateTime(2008, 2, 1), actual);
        }

        [Test]
        public void EDate_Negative2()
        {
            Object actual = XLWorkbook.EvaluateExpr("EDate(\"3/31/2008\", -1)");
            Assert.AreEqual(new DateTime(2008, 2, 29), actual);
        }

        [Test]
        public void EDate_Positive1()
        {
            Object actual = XLWorkbook.EvaluateExpr("EDate(\"3/1/2008\", 1)");
            Assert.AreEqual(new DateTime(2008, 4, 1), actual);
        }

        [Test]
        public void EDate_Positive2()
        {
            Object actual = XLWorkbook.EvaluateExpr("EDate(\"3/31/2008\", 1)");
            Assert.AreEqual(new DateTime(2008, 4, 30), actual);
        }

        [Test]
        public void EOMonth_Negative()
        {
            Object actual = XLWorkbook.EvaluateExpr("EOMonth(\"3/1/2008\", -1)");
            Assert.AreEqual(new DateTime(2008, 2, 29), actual);
        }

        [Test]
        public void EOMonth_Positive()
        {
            Object actual = XLWorkbook.EvaluateExpr("EOMonth(\"3/31/2008\", 1)");
            Assert.AreEqual(new DateTime(2008, 4, 30), actual);
        }

        [Test]
        public void Hour()
        {
            Object actual = XLWorkbook.EvaluateExpr("Hour(\"8/22/2008 3:30:45 PM\")");
            Assert.AreEqual(15, actual);
        }

        [Test]
        public void Minute()
        {
            Object actual = XLWorkbook.EvaluateExpr("Minute(\"8/22/2008 3:30:45 AM\")");
            Assert.AreEqual(30, actual);
        }

        [Test]
        public void Month()
        {
            Object actual = XLWorkbook.EvaluateExpr("Month(\"8/22/2008\")");
            Assert.AreEqual(8, actual);
        }

        [Test]
        public void IsoWeekNum()
        {
            Object actual = XLWorkbook.EvaluateExpr("ISOWEEKNUM(DATEVALUE(\"2012-3-9\"))");
            Assert.AreEqual(10, actual);

            actual = XLWorkbook.EvaluateExpr("ISOWEEKNUM(DATE(2012,12,31))");
            Assert.AreEqual(1, actual);
        }

        [Test]
        public void Networkdays_MultipleHolidaysGiven()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Date")
                .CellBelow().SetValue(new DateTime(2008, 10, 1))
                .CellBelow().SetValue(new DateTime(2009, 3, 1))
                .CellBelow().SetValue(new DateTime(2008, 11, 26))
                .CellBelow().SetValue(new DateTime(2008, 12, 4))
                .CellBelow().SetValue(new DateTime(2009, 1, 21));
            Object actual = ws.Evaluate("Networkdays(A2,A3,A4:A6)");
            Assert.AreEqual(105, actual);
        }

        [Test]
        public void Networkdays_NoHolidaysGiven()
        {
            Object actual = XLWorkbook.EvaluateExpr("Networkdays(\"10/01/2008\", \"3/01/2009\")");
            Assert.AreEqual(108, actual);
        }

        [Test]
        public void Networkdays_NegativeResult()
        {
            Object actual = XLWorkbook.EvaluateExpr("Networkdays(\"3/01/2009\", \"10/01/2008\")");
            Assert.AreEqual(-108, actual);

            actual = XLWorkbook.EvaluateExpr("Networkdays(\"2016-01-01\", \"2015-12-23\")");
            Assert.AreEqual(-8, actual);
        }

        [Test]
        public void Networkdays_OneHolidaysGiven()
        {
            Object actual = XLWorkbook.EvaluateExpr("Networkdays(\"10/01/2008\", \"3/01/2009\", \"11/26/2008\")");
            Assert.AreEqual(107, actual);
        }

        [Test]
        public void Second()
        {
            Object actual = XLWorkbook.EvaluateExpr("Second(\"8/22/2008 3:30:45 AM\")");
            Assert.AreEqual(45, actual);
        }

        [Test]
        public void Time()
        {
            Object actual = XLWorkbook.EvaluateExpr("Time(1,2,3)");
            Assert.AreEqual(new TimeSpan(1, 2, 3), actual);
        }

        [Test]
        public void TimeValue1()
        {
            Object actual = XLWorkbook.EvaluateExpr("TimeValue(\"2:24 AM\")");
            Assert.IsTrue(XLHelper.AreEqual(0.1, (double)actual));
        }

        [Test]
        public void TimeValue2()
        {
            Object actual = XLWorkbook.EvaluateExpr("TimeValue(\"22-Aug-2008 6:35 AM\")");
            Assert.IsTrue(XLHelper.AreEqual(0.27430555555555558, (double)actual));
        }

        [Test]
        public void Today()
        {
            Object actual = XLWorkbook.EvaluateExpr("Today()");
            Assert.AreEqual(DateTime.Now.Date, actual);
        }

        [Test]
        public void Weekday_1()
        {
            Object actual = XLWorkbook.EvaluateExpr("Weekday(\"2/14/2008\", 1)");
            Assert.AreEqual(5, actual);
        }

        [Test]
        public void Weekday_2()
        {
            Object actual = XLWorkbook.EvaluateExpr("Weekday(\"2/14/2008\", 2)");
            Assert.AreEqual(4, actual);
        }

        [Test]
        public void Weekday_3()
        {
            Object actual = XLWorkbook.EvaluateExpr("Weekday(\"2/14/2008\", 3)");
            Assert.AreEqual(3, actual);
        }

        [Test]
        public void Weekday_Omitted()
        {
            Object actual = XLWorkbook.EvaluateExpr("Weekday(\"2/14/2008\")");
            Assert.AreEqual(5, actual);
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
            Object actual = XLWorkbook.EvaluateExpr("Weeknum(\"3/9/2008\")");
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
            Object actual = ws.Evaluate("Workday(A2,A3,A4:A6)");
            Assert.AreEqual(new DateTime(2009, 5, 5), actual);
        }

        [Test]
        public void Workdays_NoHolidaysGiven()
        {
            Object actual = XLWorkbook.EvaluateExpr("Workday(\"10/01/2008\", 151)");
            Assert.AreEqual(new DateTime(2009, 4, 30), actual);

            actual = XLWorkbook.EvaluateExpr("Workday(\"2016-01-01\", -10)");
            Assert.AreEqual(new DateTime(2015, 12, 18), actual);
        }

        [Test]
        public void Workdays_OneHolidaysGiven()
        {
            Object actual = XLWorkbook.EvaluateExpr("Workday(\"10/01/2008\", 152, \"11/26/2008\")");
            Assert.AreEqual(new DateTime(2009, 5, 4), actual);
        }

        [Test]
        public void Year()
        {
            Object actual = XLWorkbook.EvaluateExpr("Year(\"8/22/2008\")");
            Assert.AreEqual(2008, actual);
        }

        [Test]
        public void Yearfrac_1_base0()
        {
            Object actual = XLWorkbook.EvaluateExpr("Yearfrac(\"1/1/2008\", \"3/31/2008\",0)");
            Assert.IsTrue(XLHelper.AreEqual(0.25, (double)actual));
        }

        [Test]
        public void Yearfrac_1_base1()
        {
            Object actual = XLWorkbook.EvaluateExpr("Yearfrac(\"1/1/2008\", \"3/31/2008\",1)");
            Assert.IsTrue(XLHelper.AreEqual(0.24590163934426229, (double)actual));
        }

        [Test]
        public void Yearfrac_1_base2()
        {
            Object actual = XLWorkbook.EvaluateExpr("Yearfrac(\"1/1/2008\", \"3/31/2008\",2)");
            Assert.IsTrue(XLHelper.AreEqual(0.25, (double)actual));
        }

        [Test]
        public void Yearfrac_1_base3()
        {
            Object actual = XLWorkbook.EvaluateExpr("Yearfrac(\"1/1/2008\", \"3/31/2008\",3)");
            Assert.IsTrue(XLHelper.AreEqual(0.24657534246575341, (double)actual));
        }

        [Test]
        public void Yearfrac_1_base4()
        {
            Object actual = XLWorkbook.EvaluateExpr("Yearfrac(\"1/1/2008\", \"3/31/2008\",4)");
            Assert.IsTrue(XLHelper.AreEqual(0.24722222222222223, (double)actual));
        }

        [Test]
        public void Yearfrac_2_base0()
        {
            Object actual = XLWorkbook.EvaluateExpr("Yearfrac(\"1/1/2008\", \"3/31/2013\",0)");
            Assert.IsTrue(XLHelper.AreEqual(5.25, (double)actual));
        }

        [Test]
        public void Yearfrac_2_base1()
        {
            Object actual = XLWorkbook.EvaluateExpr("Yearfrac(\"1/1/2008\", \"3/31/2013\",1)");
            Assert.IsTrue(XLHelper.AreEqual(5.24452554744526, (double)actual));
        }

        [Test]
        public void Yearfrac_2_base2()
        {
            Object actual = XLWorkbook.EvaluateExpr("Yearfrac(\"1/1/2008\", \"3/31/2013\",2)");
            Assert.IsTrue(XLHelper.AreEqual(5.32222222222222, (double)actual));
        }

        [Test]
        public void Yearfrac_2_base3()
        {
            Object actual = XLWorkbook.EvaluateExpr("Yearfrac(\"1/1/2008\", \"3/31/2013\",3)");
            Assert.IsTrue(XLHelper.AreEqual(5.24931506849315, (double)actual));
        }

        [Test]
        public void Yearfrac_2_base4()
        {
            Object actual = XLWorkbook.EvaluateExpr("Yearfrac(\"1/1/2008\", \"3/31/2013\",4)");
            Assert.IsTrue(XLHelper.AreEqual(5.24722222222222, (double)actual));
        }
    }
}
