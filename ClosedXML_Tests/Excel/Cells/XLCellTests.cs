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

            cell.Value = 5;
            cell.Value = doubleList;
            Assert.AreEqual(XLDataType.Text, cell.DataType);
            Assert.AreEqual(CultureInfo.CurrentCulture.NumberFormat.PositiveInfinitySymbol, cell.Value);

            cell.Value = 5;
            Assert.Throws<ArgumentException>(() => cell.SetValue(doubleList));
        }

        [Test]
        public void Double_NaN_is_a_string()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell("A1");
            var doubleList = new List<Double> { 0.0 / 0.0 };

            cell.Value = 5;
            cell.Value = doubleList;
            Assert.AreEqual(XLDataType.Text, cell.DataType);
            Assert.AreEqual(CultureInfo.CurrentCulture.NumberFormat.NaNSymbol, cell.Value);

            cell.Value = 5;
            Assert.Throws<ArgumentException>(() => cell.SetValue(doubleList));
        }

        [Test]
        public void InsertData1()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRange range = ws.Cell(2, 2).InsertData(new[] { "a", "b", "c" });
            Assert.AreEqual("Sheet1!B2:B4", range.ToString());
        }

        [Test]
        public void InsertData2()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRange range = ws.Cell(2, 2).InsertData(new[] { "a", "b", "c" }, false);
            Assert.AreEqual("Sheet1!B2:B4", range.ToString());
        }

        [Test]
        public void InsertData3()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRange range = ws.Cell(2, 2).InsertData(new[] { "a", "b", "c" }, true);
            Assert.AreEqual("Sheet1!B2:D2", range.ToString());
        }

        [Test]
        public void InsertData_with_Guids()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.FirstCell().InsertData(Enumerable.Range(1, 20).Select(i => new { Guid = Guid.NewGuid() }));

            Assert.AreEqual(XLDataType.Text, ws.FirstCell().DataType);
            Assert.AreEqual(Guid.NewGuid().ToString().Length, ws.FirstCell().GetString().Length);
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

            Assert.AreNotEqual(XLDataType.Number, cell.DataType);
        }

        [Test]
        public void Nan_is_not_a_number()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell("A1");
            cell.Value = "Nan";

            Assert.AreNotEqual(XLDataType.Number, cell.DataType);
        }

        [Test]
        public void TryGetValue_Boolean_Bad()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell("A1").SetValue("ABC");
            bool success = cell.TryGetValue(out bool outValue);
            Assert.IsFalse(success);
        }

        [Test]
        public void TryGetValue_Boolean_False()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell("A1").SetValue(false);
            bool success = cell.TryGetValue(out bool outValue);
            Assert.IsTrue(success);
            Assert.IsFalse(outValue);
        }

        [Test]
        public void TryGetValue_Boolean_Good()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell("A1").SetValue("True");
            bool success = cell.TryGetValue(out bool outValue);
            Assert.IsTrue(success);
            Assert.IsTrue(outValue);
        }

        [Test]
        public void TryGetValue_Boolean_True()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell("A1").SetValue(true);
            bool success = cell.TryGetValue(out bool outValue);
            Assert.IsTrue(success);
            Assert.IsTrue(outValue);
        }

        [Test]
        public void TryGetValue_DateTime_BadString()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var date = "ABC";
            bool success = ws.Cell("A1").SetValue(date).TryGetValue(out DateTime outValue);
            Assert.IsFalse(success);
        }

        [Test]
        public void TryGetValue_DateTime_BadString2()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var date = 5545454;
            ws.FirstCell().SetValue(date).DataType = XLDataType.DateTime;
            bool success = ws.FirstCell().TryGetValue(out DateTime outValue);
            Assert.IsFalse(success);
        }

        [Test]
        public void TryGetValue_RichText_Bad()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell("A1").SetValue("Anything");
            bool success = cell.TryGetValue(out IXLRichText outValue);
            Assert.IsTrue(success);
            Assert.AreEqual(cell.RichText, outValue);
            Assert.AreEqual("Anything", outValue.ToString());
        }

        [Test]
        public void TryGetValue_RichText_Good()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell("A1");
            cell.RichText.AddText("Anything");
            bool success = cell.TryGetValue(out IXLRichText outValue);
            Assert.IsTrue(success);
            Assert.AreEqual(cell.RichText, outValue);
        }

        [Test]
        public void TryGetValue_TimeSpan_BadString()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            string timeSpan = "ABC";
            bool success = ws.Cell("A1").SetValue(timeSpan).TryGetValue(out TimeSpan outValue);
            Assert.IsFalse(success);
        }

        [Test]
        public void TryGetValue_TimeSpan_Good()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var timeSpan = new TimeSpan(1, 1, 1);
            bool success = ws.Cell("A1").SetValue(timeSpan).TryGetValue(out TimeSpan outValue);
            Assert.IsTrue(success);
            Assert.AreEqual(timeSpan, outValue);
        }

        [Test]
        public void TryGetValue_TimeSpan_GoodString()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var timeSpan = new TimeSpan(1, 1, 1);
            bool success = ws.Cell("A1").SetValue(timeSpan.ToString()).TryGetValue(out TimeSpan outValue);
            Assert.IsTrue(success);
            Assert.AreEqual(timeSpan, outValue);
        }

        [Test]
        public void TryGetValue_sbyte_Bad()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell("A1").SetValue(255);
            bool success = cell.TryGetValue(out sbyte outValue);
            Assert.IsFalse(success);
        }

        [Test]
        public void TryGetValue_sbyte_Bad2()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell("A1").SetValue("255");
            bool success = cell.TryGetValue(out sbyte outValue);
            Assert.IsFalse(success);
        }

        [Test]
        public void TryGetValue_sbyte_Good()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell("A1").SetValue(5);
            bool success = cell.TryGetValue(out sbyte outValue);
            Assert.IsTrue(success);
            Assert.AreEqual(5, outValue);
        }

        [Test]
        public void TryGetValue_sbyte_Good2()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell("A1").SetValue("5");
            bool success = cell.TryGetValue(out sbyte outValue);
            Assert.IsTrue(success);
            Assert.AreEqual(5, outValue);
        }

        [Test]
        public void SetCellValueToGuid()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet1");
            var guid = Guid.NewGuid();
            ws.FirstCell().Value = guid;
            Assert.AreEqual(XLDataType.Text, ws.FirstCell().DataType);
            Assert.AreEqual(guid.ToString(), ws.FirstCell().Value);
            Assert.AreEqual(guid.ToString(), ws.FirstCell().GetString());

            guid = Guid.NewGuid();
            ws.FirstCell().SetValue(guid);
            Assert.AreEqual(XLDataType.Text, ws.FirstCell().DataType);
            Assert.AreEqual(guid.ToString(), ws.FirstCell().Value);
            Assert.AreEqual(guid.ToString(), ws.FirstCell().GetString());
        }

        [Test]
        public void SetCellValueToEnum()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet1");
            var dataType = XLDataType.Number;
            ws.FirstCell().Value = dataType;
            Assert.AreEqual(XLDataType.Text, ws.FirstCell().DataType);
            Assert.AreEqual(dataType.ToString(), ws.FirstCell().Value);
            Assert.AreEqual(dataType.ToString(), ws.FirstCell().GetString());

            dataType = XLDataType.TimeSpan;
            ws.FirstCell().SetValue(dataType);
            Assert.AreEqual(XLDataType.Text, ws.FirstCell().DataType);
            Assert.AreEqual(dataType.ToString(), ws.FirstCell().Value);
            Assert.AreEqual(dataType.ToString(), ws.FirstCell().GetString());
        }

        [Test]
        public void ValueSetToEmptyString()
        {
            string expected = String.Empty;

            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            cell.Value = new DateTime(2000, 1, 2);
            cell.Value = String.Empty;
            Assert.AreEqual(expected, cell.GetString());
            Assert.AreEqual(expected, cell.Value);

            cell.Value = new DateTime(2000, 1, 2);
            cell.SetValue(string.Empty);
            Assert.AreEqual(expected, cell.GetString());
            Assert.AreEqual(expected, cell.Value);
        }

        [Test]
        public void ValueSetToNull()
        {
            string expected = String.Empty;

            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            cell.Value = new DateTime(2000, 1, 2);
            cell.Value = null;
            Assert.AreEqual(expected, cell.GetString());
            Assert.AreEqual(expected, cell.Value);

            cell.Value = new DateTime(2000, 1, 2);
            cell.SetValue(null as string);
            Assert.AreEqual(expected, cell.GetString());
            Assert.AreEqual(expected, cell.Value);
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
        public void SetStringCellValues()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                var cell = ws.FirstCell();

                object expected;

                var date = new DateTime(2018, 4, 18);
                expected = date.ToString();
                cell.Value = expected;
                Assert.AreEqual(XLDataType.DateTime, cell.DataType);
                Assert.AreEqual(date, cell.Value);

                var b = true;
                expected = b.ToString();
                cell.Value = expected;
                Assert.AreEqual(XLDataType.Boolean, cell.DataType);
                Assert.AreEqual(b, cell.Value);

                var ts = new TimeSpan(8, 12, 4);
                expected = ts.ToString();
                cell.Value = expected;
                Assert.AreEqual(XLDataType.TimeSpan, cell.DataType);
                Assert.AreEqual(ts, cell.Value);
            }
        }

        [Test]
        public void SetStringValueTooLong()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                ws.FirstCell().Value = new DateTime(2018, 5, 15);

                ws.FirstCell().SetValue(new String('A', 32767));

                Assert.Throws<ArgumentOutOfRangeException>(() => ws.FirstCell().Value = new String('A', 32768));
                Assert.Throws<ArgumentOutOfRangeException>(() => ws.FirstCell().SetValue(new String('A', 32768)));
            }
        }

        [Test]
        public void SetDateOutOfRange()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-ZA");

            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                ws.FirstCell().Value = 5;

                var date = XLCell.BaseDate.AddDays(-1);
                ws.FirstCell().Value = date;

                // Should default to string representation using current culture's date format
                Assert.AreEqual(XLDataType.Text, ws.FirstCell().DataType);
                Assert.AreEqual(date.ToString(), ws.FirstCell().Value);

                Assert.Throws<ArgumentException>(() => ws.FirstCell().SetValue(XLCell.BaseDate.AddDays(-1)));
            }
        }

        [Test]
        public void SetCellValueWipesFormulas()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                ws.FirstCell().FormulaA1 = "=TODAY()";
                ws.FirstCell().Value = "hello world";
                Assert.IsFalse(ws.FirstCell().HasFormula);

                ws.FirstCell().FormulaA1 = "=TODAY()";
                ws.FirstCell().SetValue("hello world");
                Assert.IsFalse(ws.FirstCell().HasFormula);
            }
        }

        [Test]
        public void CellValueLineWrapping()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                ws.FirstCell().Value = "hello world";
                Assert.IsFalse(ws.FirstCell().Style.Alignment.WrapText);

                ws.FirstCell().Value = "hello\r\nworld";
                Assert.IsTrue(ws.FirstCell().Style.Alignment.WrapText);

                ws.FirstCell().Style.Alignment.WrapText = false;

                ws.FirstCell().SetValue("hello world");
                Assert.IsFalse(ws.FirstCell().Style.Alignment.WrapText);

                ws.FirstCell().SetValue("hello\r\nworld");
                Assert.IsTrue(ws.FirstCell().Style.Alignment.WrapText);
            }
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

        [Test]
        public void CanClearCellValueBySettingNullValue()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                var cell = ws.FirstCell();

                cell.Value = "Test";
                Assert.AreEqual("Test", cell.Value);
                Assert.AreEqual(XLDataType.Text, cell.DataType);

                string s = null;
                cell.SetValue(s);
                Assert.AreEqual(string.Empty, cell.Value);

                cell.Value = "Test";
                cell.Value = null;
                Assert.AreEqual(string.Empty, cell.Value);
            }
        }

        [Test]
        public void CanClearDateTimeCellValue()
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = new XLWorkbook())
                {
                    var ws = wb.AddWorksheet("Sheet1");
                    var c = ws.FirstCell();
                    c.SetValue(new DateTime(2017, 10, 08));
                    Assert.AreEqual(XLDataType.DateTime, c.DataType);
                    Assert.AreEqual(new DateTime(2017, 10, 08), c.Value);

                    wb.SaveAs(ms);
                }

                using (var wb = new XLWorkbook(ms))
                {
                    var ws = wb.Worksheets.First();
                    var c = ws.FirstCell();
                    Assert.AreEqual(XLDataType.DateTime, c.DataType);
                    Assert.AreEqual(new DateTime(2017, 10, 08), c.Value);

                    c.Clear();
                    wb.Save();
                }

                using (var wb = new XLWorkbook(ms))
                {
                    var ws = wb.Worksheets.First();
                    var c = ws.FirstCell();
                    Assert.AreEqual(XLDataType.Text, c.DataType);
                    Assert.True(c.IsEmpty());
                }
            }
        }

        [Test]
        public void CurrentRegion()
        {
            // Partially based on sample in https://github.com/ClosedXML/ClosedXML/issues/120
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                ws.Cell("B1").SetValue("x")
                    .CellBelow().SetValue("x")
                    .CellBelow().SetValue("x");

                ws.Cell("C1").SetValue("x")
                    .CellBelow().SetValue("x")
                    .CellBelow().SetValue("x");

                //Deliberately D2
                ws.Cell("D2").SetValue("x")
                    .CellBelow().SetValue("x");

                ws.Cell("G1").SetValue("x")
                    .CellBelow() // skip a cell
                    .CellBelow().SetValue("x")
                    .CellBelow().SetValue("x");

                // Deliberately H2
                ws.Cell("H2").SetValue("x")
                    .CellBelow().SetValue("x")
                    .CellBelow().SetValue("x");

                // A diagonal
                ws.Cell("E8").SetValue("x")
                    .CellBelow().CellRight().SetValue("x")
                    .CellBelow().CellRight().SetValue("x")
                    .CellBelow().CellRight().SetValue("x")
                    .CellBelow().CellRight().SetValue("x");

                Assert.AreEqual("A10:A10", ws.Cell("A10").CurrentRegion.RangeAddress.ToString());
                Assert.AreEqual("B5:B5", ws.Cell("B5").CurrentRegion.RangeAddress.ToString());
                Assert.AreEqual("P1:P1", ws.Cell("P1").CurrentRegion.RangeAddress.ToString());

                Assert.AreEqual("B1:D3", ws.Cell("D3").CurrentRegion.RangeAddress.ToString());
                Assert.AreEqual("B1:D4", ws.Cell("D4").CurrentRegion.RangeAddress.ToString());
                Assert.AreEqual("B1:E4", ws.Cell("E4").CurrentRegion.RangeAddress.ToString());

                foreach (var c in ws.Range("B1:D3").Cells())
                {
                    Assert.AreEqual("B1:D3", c.CurrentRegion.RangeAddress.ToString());
                }

                foreach (var c in ws.Range("A1:A3").Cells())
                {
                    Assert.AreEqual("A1:D3", c.CurrentRegion.RangeAddress.ToString());
                }

                Assert.AreEqual("A1:D4", ws.Cell("A4").CurrentRegion.RangeAddress.ToString());

                foreach (var c in ws.Range("E1:E3").Cells())
                {
                    Assert.AreEqual("B1:E3", c.CurrentRegion.RangeAddress.ToString());
                }

                Assert.AreEqual("B1:E4", ws.Cell("E4").CurrentRegion.RangeAddress.ToString());

                //// SECOND REGION
                foreach (var c in ws.Range("F1:F4").Cells())
                {
                    Assert.AreEqual("F1:H4", c.CurrentRegion.RangeAddress.ToString());
                }

                Assert.AreEqual("F1:H5", ws.Cell("F5").CurrentRegion.RangeAddress.ToString());

                //// DIAGONAL
                Assert.AreEqual("E8:I12", ws.Cell("E8").CurrentRegion.RangeAddress.ToString());
                Assert.AreEqual("E8:I12", ws.Cell("F9").CurrentRegion.RangeAddress.ToString());
                Assert.AreEqual("E8:I12", ws.Cell("G10").CurrentRegion.RangeAddress.ToString());
                Assert.AreEqual("E8:I12", ws.Cell("H11").CurrentRegion.RangeAddress.ToString());
                Assert.AreEqual("E8:I12", ws.Cell("I12").CurrentRegion.RangeAddress.ToString());

                Assert.AreEqual("E8:I12", ws.Cell("G9").CurrentRegion.RangeAddress.ToString());
                Assert.AreEqual("E8:I12", ws.Cell("F10").CurrentRegion.RangeAddress.ToString());

                Assert.AreEqual("D7:I12", ws.Cell("D7").CurrentRegion.RangeAddress.ToString());
                Assert.AreEqual("E8:J13", ws.Cell("J13").CurrentRegion.RangeAddress.ToString());
            }
        }

        // https://github.com/ClosedXML/ClosedXML/issues/630
        [Test]
        public void ConsiderEmptyValueAsNumericInSumFormula()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                ws.Cell("A1").SetValue("Empty");
                ws.Cell("A2").SetValue("Numeric");
                ws.Cell("A3").SetValue("Copy of numeric");

                ws.Cell("B2").SetFormulaA1("=B1");
                ws.Cell("B3").SetFormulaA1("=B2");

                ws.Cell("C2").SetFormulaA1("=SUM(C1)");
                ws.Cell("C3").SetFormulaA1("=C2");

                object b1 = ws.Cell("B1").Value;
                object b2 = ws.Cell("B2").Value;
                object b3 = ws.Cell("B3").Value;

                Assert.AreEqual("", b1);
                Assert.AreEqual(0, b2);
                Assert.AreEqual(0, b3);

                object c1 = ws.Cell("C1").Value;
                object c2 = ws.Cell("C2").Value;
                object c3 = ws.Cell("C3").Value;

                Assert.AreEqual("", c1);
                Assert.AreEqual(0, c2);
                Assert.AreEqual(0, c3);
            }
        }

        [Test]
        public void SetFormulaA1AffectsR1C1()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                var cell = ws.Cell(1, 1);
                cell.FormulaR1C1 = "R[1]C";

                cell.FormulaA1 = "B2";

                Assert.AreEqual("R[1]C[1]", cell.FormulaR1C1);
            }
        }

        [Test]
        public void SetFormulaR1C1AffectsA1()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                var cell = ws.Cell(1, 1);
                cell.FormulaA1 = "A2";

                cell.FormulaR1C1 = "R[1]C[1]";

                Assert.AreEqual("B2", cell.FormulaA1);
            }
        }

        [Test]
        public void FormulaWithCircularReferenceFails()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                var A1 = ws.Cell("A1");
                var A2 = ws.Cell("A2");
                A1.FormulaA1 = "A2 + 1";
                A2.FormulaA1 = "A1 + 1";

                Assert.Throws<InvalidOperationException>(() =>
                {
                    var _ = A1.Value;
                });
                Assert.Throws<InvalidOperationException>(() =>
                {
                    var _ = A2.Value;
                });
            }
        }

        [Test]
        public void InvalidFormulaShiftProducesREF()
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = new XLWorkbook())
                {
                    var ws = wb.Worksheets.Add("Sheet1");
                    ws.Cell("A1").Value = 1;
                    ws.Cell("B1").Value = 2;
                    ws.Cell("B2").FormulaA1 = "=A1+B1";

                    Assert.AreEqual(3, ws.Cell("B2").Value);

                    ws.Range("A2").Value = ws.Range("B2");
                    var fA2 = ws.Cell("A2").FormulaA1;

                    wb.SaveAs(ms);

                    Assert.AreEqual("#REF!+A1", fA2);
                }

                using (var wb2 = new XLWorkbook(ms))
                {
                    var fA2 = wb2.Worksheets.First().Cell("A2").FormulaA1;
                    Assert.AreEqual("#REF!+A1", fA2);
                }
            }
        }

        public void FormulaWithCircularReferenceFails2()
        {
            var cell = new XLWorkbook().Worksheets.Add("Sheet1").FirstCell();
            cell.FormulaA1 = "A1";
            Assert.Throws<InvalidOperationException>(() =>
            {
                var _ = cell.Value;
            });
        }

        [Test]
        public void TryGetValueFormulaEvaluation()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                var A1 = ws.Cell("A1");
                var A2 = ws.Cell("A2");
                var A3 = ws.Cell("A3");
                A1.FormulaA1 = "A2 + 1";
                A2.FormulaA1 = "A1 + 1";

                Assert.IsFalse(A1.TryGetValue(out String _));
                Assert.IsFalse(A2.TryGetValue(out String _));
                Assert.IsTrue(A3.TryGetValue(out String _));
            }
        }
    }
}
