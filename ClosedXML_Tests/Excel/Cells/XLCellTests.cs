using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;

namespace ClosedXML_Tests
{
    [TestFixture]
    public class XLCellTests
    {
        [Test]
        public void CellsUsed()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Cell(1, 1);
            ws.Cell(2, 2);
            int count = ws.Range("A1:B2").CellsUsed().Count();
            Assert.AreEqual(0, count);
        }

        [Test]
        public void CellsUsedIncludeStyles1()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Row(3).Style.Fill.BackgroundColor = XLColor.Red;
            ws.Column(3).Style.Fill.BackgroundColor = XLColor.Red;
            ws.Cell(2, 2).Value = "ASDF";
            var range = ws.RangeUsed(true).RangeAddress.ToString();
            Assert.AreEqual("B2:C3", range);
        }

        [Test]
        public void CellsUsedIncludeStyles2()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Row(2).Style.Fill.BackgroundColor = XLColor.Red;
            ws.Column(2).Style.Fill.BackgroundColor = XLColor.Red;
            ws.Cell(3, 3).Value = "ASDF";
            var range = ws.RangeUsed(true).RangeAddress.ToString();
            Assert.AreEqual("B2:C3", range);
        }

        [Test]
        public void CellsUsedIncludeStyles3()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var range = ws.RangeUsed(true);
            Assert.AreEqual(null, range);
        }

        [Test]
        public void Double_Infinity_is_a_string()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell("A1");
            var doubleList = new List<Double> { 1.0 / 0.0 };

            cell.Value = doubleList.AsEnumerable();
            Assert.AreNotEqual(XLCellValues.Number, cell.DataType);
        }

        [Test]
        public void Double_NaN_is_a_string()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell("A1");
            var doubleList = new List<Double> { 0.0 / 0.0 };

            cell.Value = doubleList.AsEnumerable();
            Assert.AreNotEqual(XLCellValues.Number, cell.DataType);
        }

        [Test]
        public void InsertData1()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRange range = ws.Cell(2, 2).InsertData(new[] { "a", "b", "c" });
            Assert.AreEqual("Sheet1!B2:B4", range.ToString());
        }

        [Test]
        public void IsEmpty1()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            bool actual = cell.IsEmpty();
            bool expected = true;
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void IsEmpty2()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            bool actual = cell.IsEmpty(true);
            bool expected = true;
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void IsEmpty3()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            cell.Style.Fill.BackgroundColor = XLColor.Red;
            bool actual = cell.IsEmpty();
            bool expected = true;
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void IsEmpty4()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            cell.Style.Fill.BackgroundColor = XLColor.Red;
            bool actual = cell.IsEmpty(false);
            bool expected = true;
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void IsEmpty5()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            cell.Style.Fill.BackgroundColor = XLColor.Red;
            bool actual = cell.IsEmpty(true);
            bool expected = false;
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void IsEmpty6()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            cell.Value = "X";
            bool actual = cell.IsEmpty();
            bool expected = false;
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void NaN_is_not_a_number()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell("A1");
            cell.Value = "NaN";

            Assert.AreNotEqual(XLCellValues.Number, cell.DataType);
        }

        [Test]
        public void Nan_is_not_a_number()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell("A1");
            cell.Value = "Nan";

            Assert.AreNotEqual(XLCellValues.Number, cell.DataType);
        }

        [Test]
        public void TryGetValue_Boolean_Bad()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            Boolean outValue;
            IXLCell cell = ws.Cell("A1").SetValue("ABC");
            bool success = cell.TryGetValue(out outValue);
            Assert.IsFalse(success);
        }

        [Test]
        public void TryGetValue_Boolean_False()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            Boolean outValue;
            IXLCell cell = ws.Cell("A1").SetValue(false);
            bool success = cell.TryGetValue(out outValue);
            Assert.IsTrue(success);
            Assert.IsFalse(outValue);
        }

        [Test]
        public void TryGetValue_Boolean_Good()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            Boolean outValue;
            IXLCell cell = ws.Cell("A1").SetValue("True");
            bool success = cell.TryGetValue(out outValue);
            Assert.IsTrue(success);
            Assert.IsTrue(outValue);
        }

        [Test]
        public void TryGetValue_Boolean_True()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            Boolean outValue;
            IXLCell cell = ws.Cell("A1").SetValue(true);
            bool success = cell.TryGetValue(out outValue);
            Assert.IsTrue(success);
            Assert.IsTrue(outValue);
        }

        [Test]
        public void TryGetValue_DateTime_BadString()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            DateTime outValue;
            var date = "ABC";
            bool success = ws.Cell("A1").SetValue(date).TryGetValue(out outValue);
            Assert.IsFalse(success);
        }


        [Test]
        public void TryGetValue_DateTime_BadString2()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            DateTime outValue;
            var date = 5545454;
            ws.FirstCell().SetValue(date).DataType = XLCellValues.DateTime;
            bool success = ws.FirstCell().TryGetValue(out outValue);
            Assert.IsFalse(success);
        }

        [Test]
        public void TryGetValue_RichText_Bad()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText outValue;
            IXLCell cell = ws.Cell("A1").SetValue("Anything");
            bool success = cell.TryGetValue(out outValue);
            Assert.IsTrue(success);
            Assert.AreEqual(cell.RichText, outValue);
            Assert.AreEqual("Anything", outValue.ToString());
        }

        [Test]
        public void TryGetValue_RichText_Good()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRichText outValue;
            IXLCell cell = ws.Cell("A1");
            cell.RichText.AddText("Anything");
            bool success = cell.TryGetValue(out outValue);
            Assert.IsTrue(success);
            Assert.AreEqual(cell.RichText, outValue);
        }

        [Test]
        public void TryGetValue_TimeSpan_BadString()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            TimeSpan outValue;
            string timeSpan = "ABC";
            bool success = ws.Cell("A1").SetValue(timeSpan).TryGetValue(out outValue);
            Assert.IsFalse(success);
        }

        [Test]
        public void TryGetValue_TimeSpan_Good()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            TimeSpan outValue;
            var timeSpan = new TimeSpan(1, 1, 1);
            bool success = ws.Cell("A1").SetValue(timeSpan).TryGetValue(out outValue);
            Assert.IsTrue(success);
            Assert.AreEqual(timeSpan, outValue);
        }

        [Test]
        public void TryGetValue_TimeSpan_GoodString()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            TimeSpan outValue;
            var timeSpan = new TimeSpan(1, 1, 1);
            bool success = ws.Cell("A1").SetValue(timeSpan.ToString()).TryGetValue(out outValue);
            Assert.IsTrue(success);
            Assert.AreEqual(timeSpan, outValue);
        }

        [Test]
        public void TryGetValue_sbyte_Bad()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            sbyte outValue;
            IXLCell cell = ws.Cell("A1").SetValue(255);
            bool success = cell.TryGetValue(out outValue);
            Assert.IsFalse(success);
        }

        [Test]
        public void TryGetValue_sbyte_Bad2()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            sbyte outValue;
            IXLCell cell = ws.Cell("A1").SetValue("255");
            bool success = cell.TryGetValue(out outValue);
            Assert.IsFalse(success);
        }

        [Test]
        public void TryGetValue_sbyte_Good()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            sbyte outValue;
            IXLCell cell = ws.Cell("A1").SetValue(5);
            bool success = cell.TryGetValue(out outValue);
            Assert.IsTrue(success);
            Assert.AreEqual(5, outValue);
        }

        [Test]
        public void TryGetValue_sbyte_Good2()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            sbyte outValue;
            IXLCell cell = ws.Cell("A1").SetValue("5");
            bool success = cell.TryGetValue(out outValue);
            Assert.IsTrue(success);
            Assert.AreEqual(5, outValue);
        }

        [Test]
        public void ValueSetToEmptyString()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            cell.Value = new DateTime(2000, 1, 2);
            cell.Value = String.Empty;
            string actual = cell.GetString();
            string expected = String.Empty;
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ValueSetToNull()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            cell.Value = new DateTime(2000, 1, 2);
            cell.Value = null;
            string actual = cell.GetString();
            string expected = String.Empty;
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ValueSetDateWithShortUserDateFormat()
        {
            // For this test to make sense, user's local date format should be dd/MM/yy (note without the 2 century digits)
            // What happened previously was that the century digits got lost in .ToString() conversion and wrong century was sometimes returned.
            var ci = new CultureInfo(CultureInfo.InvariantCulture.LCID);
            ci.DateTimeFormat.ShortDatePattern = "dd/MM/yy";
            Thread.CurrentThread.CurrentCulture = ci;
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            var expected = DateTime.Today.AddYears(20);
            cell.Value = expected;
            var actual = (DateTime)cell.Value;
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void TestInvalidXmlCharacters()
        {
            byte[] data;

            using (var stream = new MemoryStream())
            {
                var wb = new XLWorkbook();
                wb.AddWorksheet("Sheet1").FirstCell().SetValue("\u0018");
                wb.SaveAs(stream);
                data = stream.ToArray();
            }

            using (var stream = new MemoryStream(data))
            {
                var wb = new XLWorkbook(stream);
                Assert.AreEqual("\u0018", wb.Worksheets.First().FirstCell().Value);
            }
        }
    }
}
