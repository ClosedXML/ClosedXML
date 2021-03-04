using ClosedXML.Excel;
using MoreLinq;
using System;
using System.Linq;

// TODO: Add example to Wiki

namespace ClosedXML.Examples.Tables
{
    public class ResizingTables : IXLExample
    {
        public void Create(string filePath)
        {
            using (var wb = new XLWorkbook())
            {
                var ws1 = wb.AddWorksheet("Sheet1");

                var data1 = Enumerable.Range(1, 10)
                    .Select(i =>
                    new
                    {
                        Index = i,
                        Character = Convert.ToChar(64 + i),
                        String = new String('a', i),
                        Integer = 64 + i
                    });

                var table1 = ws1.Cell("B2").InsertTable(data1, true)
                    .SetShowHeaderRow()
                    .SetShowTotalsRow();

                table1.Fields.First().TotalsRowLabel = "Sum of Integer";
                table1.Fields.Last().TotalsRowFunction = XLTotalsRowFunction.Sum;

                var ws2 = ws1.CopyTo("Sheet2");
                var table2 = ws2.Tables.First();
                table2.Resize(table2.FirstCell(), table2.LastCell().CellLeft().CellAbove(3));

                var ws3 = ws2.CopyTo("Sheet3");
                var table3 = ws3.Tables.First();
                table3.Resize(table3.FirstCell().CellLeft(), table3.LastCell().CellRight().CellBelow(1));

                ////See #1492
                var ws4 = ws1.CopyTo("Sheet4");
                var table4 = ws4.Tables.First();
                table4.Field("String").Column.InsertColumnsAfter(1, true);

                wb.Worksheets.ForEach(ws => ws.Columns().AdjustToContents());
                wb.SaveAs(filePath);
            }
        }
    }
}
