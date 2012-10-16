using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClosedXML_Tests.Excel.DataValidations
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    [TestClass]
    public class DateAndTimeTests
    {
        [TestMethod]
        public void Date()
        {
            Object actual = XLWorkbook.EvaluateExpr("Date(2008, 1, 1)");
            Assert.AreEqual(39448, actual);
        }

        [TestMethod]
        public void Datevalue()
        {
            Object actual = XLWorkbook.EvaluateExpr("DateValue(\"8/22/2008\")");
            Assert.AreEqual(39682, actual);
        }

        [TestMethod]
        public void Day()
        {
            Object actual = XLWorkbook.EvaluateExpr("Day(\"8/22/2008\")");
            Assert.AreEqual(22, actual);
        }

        [TestMethod]
        public void Month()
        {
            Object actual = XLWorkbook.EvaluateExpr("Month(\"8/22/2008\")");
            Assert.AreEqual(8, actual);
        }

        [TestMethod]
        public void Year()
        {
            Object actual = XLWorkbook.EvaluateExpr("Year(\"8/22/2008\")");
            Assert.AreEqual(2008, actual);
        }

        [TestMethod]
        public void Second()
        {
            Object actual = XLWorkbook.EvaluateExpr("Second(\"8/22/2008 3:30:45 AM\")");
            Assert.AreEqual(45, actual);
        }

        [TestMethod]
        public void Minute()
        {
            Object actual = XLWorkbook.EvaluateExpr("Minute(\"8/22/2008 3:30:45 AM\")");
            Assert.AreEqual(30, actual);
        }

        [TestMethod]
        public void Hour()
        {
            Object actual = XLWorkbook.EvaluateExpr("Hour(\"8/22/2008 3:30:45 PM\")");
            Assert.AreEqual(15, actual);
        }

        [TestMethod]
        public void Time()
        {
            Object actual = XLWorkbook.EvaluateExpr("Time(1,2,3)");
            Assert.AreEqual(new TimeSpan(1, 2, 3), actual);
        }

        [TestMethod]
        public void TimeValue1()
        {
            Object actual = XLWorkbook.EvaluateExpr("TimeValue(\"2:24 AM\")");
            Assert.IsTrue(XLHelper.AreEqual(0.1, (double)actual));
        }

        [TestMethod]
        public void TimeValue2()
        {
            Object actual = XLWorkbook.EvaluateExpr("TimeValue(\"22-Aug-2008 6:35 AM\")");
            Assert.IsTrue(XLHelper.AreEqual(0.27430555555555558, (double)actual));
        }

        [TestMethod]
        public void Today()
        {
            Object actual = XLWorkbook.EvaluateExpr("Today()");
            Assert.AreEqual(DateTime.Now.Date, actual);
        }
    }
}
