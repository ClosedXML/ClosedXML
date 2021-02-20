using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Data;
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
            var range = ws.RangeUsed(XLCellsUsedOptions.All).RangeAddress.ToString();
            Assert.AreEqual("B2:C3", range);
        }

        [Test]
        public void CellsUsedIncludeStyles2()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Row(2).Style.Fill.BackgroundColor = XLColor.Red;
            ws.Column(2).Style.Fill.BackgroundColor = XLColor.Red;
            ws.Cell(3, 3).Value = "ASDF";
            var range = ws.RangeUsed(XLCellsUsedOptions.All).RangeAddress.ToString();
            Assert.AreEqual("B2:C3", range);
        }

        [Test]
        public void CellsUsedIncludeStyles3()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var range = ws.RangeUsed(XLCellsUsedOptions.All);
            Assert.AreEqual(null, range);
        }

        [Test]
        public void CellUsedIncludesSparklines()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Range("C3:E4").Value = 1;
            ws.SparklineGroups.Add("B2", "C3:E3");
            ws.SparklineGroups.Add("F5", "C4:E4");

            var range = ws.RangeUsed(true).RangeAddress.ToString();
            Assert.AreEqual("B2:F5", range);
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
            cell.SetValue(doubleList);
            Assert.AreEqual(XLDataType.Text, cell.DataType);
            Assert.AreEqual(CultureInfo.CurrentCulture.NumberFormat.PositiveInfinitySymbol, cell.Value);
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
            cell.SetValue(doubleList);
            Assert.AreEqual(XLDataType.Text, cell.DataType);
            Assert.AreEqual(CultureInfo.CurrentCulture.NumberFormat.NaNSymbol, cell.Value);
        }

        [Test]
        public void GetValue_Nullable()
        {
            var backupCulture = Thread.CurrentThread.CurrentCulture;

            // Set thread culture to French, which should format numbers using a space as thousands separator
            try
            {
                var culture = CultureInfo.CreateSpecificCulture ("fr-FR");
                // but use a period instead of a comma as for decimal separator
                culture.NumberFormat.CurrencyDecimalSeparator = ".";
                Thread.CurrentThread.CurrentCulture = culture;

                var cell = new XLWorkbook ().AddWorksheet ().FirstCell ();

                Assert.IsNull (cell.Clear ().GetValue<double?> ());
                Assert.AreEqual (1.5, cell.SetValue (1.5).GetValue<double?> ());
                Assert.AreEqual (2, cell.SetValue (1.5).GetValue<int?> ());
                Assert.AreEqual (2.5, cell.SetValue ("2.5").GetValue<double?> ());
                Assert.Throws<FormatException> (() => cell.SetValue ("text").GetValue<double?> ());

            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = backupCulture;
            }
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
        public void InsertData_with_Nulls()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");

            var table = new DataTable();
            table.TableName = "Patients";
            table.Columns.Add("Dosage", typeof(int));
            table.Columns.Add("Drug", typeof(string));
            table.Columns.Add("Patient", typeof(string));
            table.Columns.Add("Date", typeof(DateTime));

            table.Rows.Add(25, "Indocin", "David", new DateTime(2000, 1, 1));
            table.Rows.Add(50, "Enebrel", "Sam", new DateTime(2000, 1, 2));
            table.Rows.Add(10, "Hydralazine", "Christoff", new DateTime(2000, 1, 3));
            table.Rows.Add(21, "Combivent", DBNull.Value, new DateTime(2000, 1, 4));
            table.Rows.Add(100, "Dilantin", "Melanie", DBNull.Value);

            ws.FirstCell().InsertData(table);

            Assert.AreEqual(25, ws.Cell("A1").Value);
            Assert.AreEqual("", ws.Cell("C4").Value);
            Assert.AreEqual("", ws.Cell("D5").Value);
        }

        [Test]
        public void InsertData_with_Nulls_IEnumerable()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");

            var dateTimeList = new List<DateTime?>()
            {
                new DateTime(2000, 1, 1),
                new DateTime(2000, 1, 2),
                new DateTime(2000, 1, 3),
                new DateTime(2000, 1, 4),
                null
            };

            ws.FirstCell().InsertData(dateTimeList);

            Assert.AreEqual(new DateTime(2000, 1, 1), ws.Cell("A1").GetDateTime());
            Assert.AreEqual(String.Empty, ws.Cell("A5").Value);
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
            bool actual = cell.IsEmpty(XLCellsUsedOptions.All);
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
            bool actual = cell.IsEmpty(XLCellsUsedOptions.AllContents);
            bool expected = true;
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void IsEmpty5()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            cell.Style.Fill.BackgroundColor = XLColor.Red;
            bool actual = cell.IsEmpty(XLCellsUsedOptions.All);
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
        public void TryGetValue_DateTime_Good()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var date = "2018-01-01";
            bool success = ws.Cell("A1").SetValue(date).TryGetValue(out DateTime outValue);
            Assert.IsTrue(success);
            Assert.AreEqual(new DateTime(2018, 1, 1), outValue);
        }

        [Test]
        public void TryGetValue_DateTime_Good2()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            bool success = ws.Cell("A1").SetFormulaA1("=TODAY() + 10").TryGetValue(out DateTime outValue);
            Assert.IsTrue(success);
            Assert.AreEqual(DateTime.Today.AddDays(10), outValue);
        }

        [Test]
        public void TryGetValue_DateTime_BadButFormulaGood()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            bool success = ws.Cell("A1").SetFormulaA1("=\"44\"&\"020\"").TryGetValue(out DateTime outValue);
            Assert.IsFalse(success);

            ws.Cell("B1").SetFormulaA1("=A1+1");

            success = ws.Cell("B1").TryGetValue(out outValue);
            Assert.IsTrue(success);
            Assert.AreEqual(new DateTime(2020, 07, 09), outValue);
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
        public void TryGetValue_Enum_Good()
        {
            var ws = new XLWorkbook().AddWorksheet();
            Assert.IsTrue(ws.FirstCell().SetValue(NumberStyles.AllowCurrencySymbol).TryGetValue(out NumberStyles value));
            Assert.AreEqual(NumberStyles.AllowCurrencySymbol, value);

            // Nullable alternative
            Assert.IsTrue(ws.FirstCell().SetValue(NumberStyles.AllowCurrencySymbol).TryGetValue(out NumberStyles? value2));
            Assert.AreEqual(NumberStyles.AllowCurrencySymbol, value2);
        }

        [Test]
        public void TryGetValue_Enum_BadString()
        {
            var ws = new XLWorkbook().AddWorksheet();
            Assert.IsFalse(ws.FirstCell().SetValue("ABC").TryGetValue(out NumberStyles value));
            Assert.IsFalse(ws.FirstCell().SetValue("ABC").TryGetValue(out NumberStyles? value2));
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
        public void TryGetValue_TimeSpan_Good_Large()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var timeSpan = TimeSpan.FromMilliseconds((double)int.MaxValue + 1);
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
        public void TryGetValue_decimal_Good()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var cell = ws.Cell("A1").SetValue("5");
            bool success = cell.TryGetValue(out decimal outValue);
            Assert.IsTrue(success);
            Assert.AreEqual(5, outValue);
        }

        [Test]
        public void TryGetValue_decimal_Good2()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");

            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var cell = ws.Cell("A1").SetValue("1.60000001869776E-06");
            bool success = cell.TryGetValue(out decimal outValue);
            Assert.IsTrue(success);
            Assert.AreEqual(1.60000001869776E-06, outValue);
        }

        [Test]
        public void TryGetValue_Hyperlink()
        {
            using (var wb = new XLWorkbook())
            {
                var ws1 = wb.Worksheets.Add("Sheet1");
                var ws2 = wb.Worksheets.Add("Sheet2");

                var targetCell = ws2.Cell("A1");

                var linkCell1 = ws1.Cell("A1");
                linkCell1.Value = "Link to IXLCell";
                linkCell1.Hyperlink = new XLHyperlink(targetCell);

                var success = linkCell1.TryGetValue(out XLHyperlink hyperlink);
                Assert.IsTrue(success);
                Assert.AreEqual("Sheet2!A1", hyperlink.InternalAddress);
            }
        }

        [Test]
        public void TryGetValue_Unicode_String()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");

            Boolean success;
            String outValue;

            success = ws.Cell("A1")
                .SetValue("Site_x0020_Column_x0020_Test")
                .TryGetValue(out outValue);
            Assert.IsTrue(success);
            Assert.AreEqual("Site Column Test", outValue);

            success = ws.Cell("A1")
                .SetValue("Site_x005F_x0020_Column_x005F_x0020_Test")
                .TryGetValue(out outValue);

            Assert.IsTrue(success);
            Assert.AreEqual("Site_x005F_x0020_Column_x005F_x0020_Test", outValue);
        }

        [Test]
        public void TryGetValue_Nullable()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            ws.Cell("A1").Clear();
            ws.Cell("A2").SetValue(1.5);
            ws.Cell("A3").SetValue("2.5");
            ws.Cell("A4").SetValue("text");

            foreach (var cell in ws.Range("A1:A3").Cells())
            {
                Assert.IsTrue(cell.TryGetValue(out double? value));
            }

            Assert.IsFalse(ws.Cell("A4").TryGetValue(out double? _));
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
        public void SetCellValueToRange()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet1");

            ws.Cell("A1").SetValue(2)
                .CellRight().SetValue(3)
                .CellRight().SetValue(5)
                .CellRight().SetValue(7);

            var range = ws.Range("1:1");

            ws.Cell("B2").Value = range;

            Assert.AreEqual(2, ws.Cell("B2").Value);
            Assert.AreEqual(3, ws.Cell("C2").Value);
            Assert.AreEqual(5, ws.Cell("D2").Value);
            Assert.AreEqual(7, ws.Cell("E2").Value);
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
        public void ClearCellRemovesSparkline()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.SparklineGroups.Add("B1:B3", "C1:E3");

            ws.Cell("B1").Clear(XLClearOptions.All);
            ws.Cell("B2").Clear(XLClearOptions.Sparklines);

            Assert.AreEqual(1, ws.SparklineGroups.Single().Count());
            Assert.IsFalse(ws.Cell("B1").HasSparkline);
            Assert.IsFalse(ws.Cell("B2").HasSparkline);
            Assert.IsTrue(ws.Cell("B3").HasSparkline);
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

        [Test]
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

        [Test]
        public void SetValue_IEnumerable()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            object[] values = { "Text", 45, DateTime.Today, true, "More text" };

            ws.FirstCell().SetValue(values);

            Assert.AreEqual("Text", ws.FirstCell().GetString());
            Assert.AreEqual(45, ws.Cell("A2").GetDouble());
            Assert.AreEqual(DateTime.Today, ws.Cell("A3").GetDateTime());
            Assert.AreEqual(true, ws.Cell("A4").GetBoolean());
            Assert.AreEqual("More text", ws.Cell("A5").GetString());
            Assert.IsTrue(ws.Cell("A6").IsEmpty());
        }

        [Test]
        public void ToStringNoFormatString()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var c = ws.FirstCell().CellBelow(2).CellRight(3);

            Assert.AreEqual("D3", c.ToString());
        }

        [Test]
        [TestCase("D3", "A")]
        [TestCase("YEAR(DATE(2018, 1, 1))", "F")]
        [TestCase("YEAR(DATE(2018, 1, 1))", "f")]
        [TestCase("0000.00", "NF")]
        [TestCase("0000.00", "nf")]
        [TestCase("FFFF0000", "fg")]
        [TestCase("Color Theme: Accent5, Tint: 0", "BG")]
        [TestCase("2018.00", "v")]
        public void ToStringFormatString(string expected, string format)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var c = ws.FirstCell().CellBelow(2).CellRight(3);

            var formula = "YEAR(DATE(2018, 1, 1))";
            c.FormulaA1 = formula;

            var numberFormat = "0000.00";
            c.Style.NumberFormat.Format = numberFormat;

            c.Style.Font.FontColor = XLColor.Red;
            c.Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent5);

            Assert.AreEqual(expected, c.ToString(format));

            Assert.Throws<FormatException>(() => c.ToString("dummy"));
        }

        [Test]
        public void ToStringInvalidFormat()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var c = ws.FirstCell();

            Assert.Throws<FormatException>(() => c.ToString("dummy"));
        }
    }
}
