using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Data;

namespace ClosedXML_Tests.Excel
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
                ws.Cell("A1").InsertData(table.AsEnumerable(), true);
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
                    .SetValue(41573.875)
                    .SetDataType(XLDataType.DateTime);

                c.Style.NumberFormat.SetFormat("m/d/yy\\ h:mm;@");

                Assert.AreEqual("10/26/13 21:00", c.GetFormattedString());
            }
        }
    }
}
