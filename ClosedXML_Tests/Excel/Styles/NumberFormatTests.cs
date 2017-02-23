using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NUnit.Framework;
using ClosedXML.Excel;
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
                ws.Column(1).Style.NumberFormat.Format = "yy-MM-dd";

                var table = new DataTable();
                table.Columns.Add("Date", typeof(DateTime));

                for (int i = 0; i <10; i++)
                {
                    table.Rows.Add(new DateTime(2017, 1, 1).AddMonths(i));
                }

                ws.Cell("B1").Value = table.Columns[0].DataType;

                ws.Cell("A1").InsertData(table.AsEnumerable());

                Assert.AreEqual("yy-MM-dd", ws.Cell("A5").Style.DateFormat.Format);
            }
        }
    }
}
