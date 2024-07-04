using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;

namespace ClosedXML.Tests.Excel
{
    public class NumberFormatTests
    {
        [Test]
        public void PreserveCellFormat()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                var table = new DataTable();
                table.Columns.Add("Date", typeof(DateTime));

                for (int i = 0; i < 10; i++)
                {
                    table.Rows.Add(new DateTime(2017, 1, 1).AddMonths(i));
                }

                ws.Column(1).Style.NumberFormat.Format = "yy-MM-dd";
                ws.Cell("A1").InsertData(table);
                Assert.AreEqual("yy-MM-dd", ws.Cell("A5").Style.DateFormat.Format);

                ws.Row(1).Style.NumberFormat.Format = "yy-MM-dd";
                ws.Cell("A1").InsertData(table.Rows, true);
                Assert.AreEqual("yy-MM-dd", ws.Cell("E1").Style.DateFormat.Format);
            }
        }

        [Test]
        public void TestExcelNumberFormats()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                var c = ws.FirstCell()
                    .SetValue((41573.875));

                c.Style.NumberFormat.SetFormat("m/d/yy\\ h:mm;@");

                Assert.AreEqual("10/26/13 21:00", c.GetFormattedString());
            }
        }

        [Test]
        [SetCulture("en-US")]
        public void Cell_value_is_formatted_by_current_culture_unless_specified_otherwise()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var cell = ws.Cell("A1").SetValue(10000.5);

            var currentCultureFormat = cell.GetFormattedString();
            Assert.AreEqual("10000.5", currentCultureFormat);

            var czechCultureFormat = cell.GetFormattedString(CultureInfo.GetCultureInfo("cs-CZ"));
            Assert.AreEqual("10000,5", czechCultureFormat);
        }

        [Test]
        public void ReadAndWriteColumnNumberFormat()
        {
            using (var memoryStream = new MemoryStream())
            {
                using (var wb = new XLWorkbook())
                {
                    var ws = wb.AddWorksheet();
                    var sourceColumn = ws.Column(1);
                    sourceColumn.Style.NumberFormat.Format = "0.000";
                    wb.SaveAs(memoryStream);
                }

                memoryStream.Position = 0;

                using (var wb = new XLWorkbook(memoryStream))
                {
                    var column = wb.Worksheets.Single().Column(1);
                    Assert.AreEqual("0.000", column.Style.NumberFormat.Format);
                }
            }
        }

        [Test]
        public void XLNumberFormatKey_GetHashCode_IsCaseSensitive()
        {
            var numberFormatKey1 = XLNumberFormatKey.ForFormat("MM");
            var numberFormatKey2 = XLNumberFormatKey.ForFormat("mm");

            Assert.AreNotEqual(numberFormatKey1.GetHashCode(), numberFormatKey2.GetHashCode());
        }

        [Test]
        public void XLNumberFormatKey_Equals_IsCaseSensitive()
        {
            var numberFormatKey1 = XLNumberFormatKey.ForFormat("MM");
            var numberFormatKey2 = XLNumberFormatKey.ForFormat("mm");

            Assert.IsFalse(numberFormatKey1.Equals(numberFormatKey2));
        }

        [Test]
        public void AddCustomNumberFormatsToFileWithNonSequentialNumberFormatIds()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\NumberFormats\NonSequentialNumberFormatsIds-Input.xlsx")))
            {
                TestHelper.CreateAndCompare(() =>
                {
                    var wb = new XLWorkbook(stream);

                    var ws = wb.Worksheet("Sheet1");

                    var format = "\"P\" #,##0.00; \"N\" #,##0.00;0;@";
                    ws.Cell(5, 1).Value = 1.2;
                    ws.Cell(5, 1).Style.NumberFormat.Format = format;
                    ws.Cell(5, 2).Value = -1.2;
                    ws.Cell(5, 2).Style.NumberFormat.Format = format;

                    return wb;
                }, @"Other\NumberFormats\NonSequentialNumberFormatsIds-Output.xlsx");
            }
        }
    }
}
