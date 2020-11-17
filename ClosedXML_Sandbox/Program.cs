using ClosedXML.Excel;
using System;
using System.Linq;

namespace ClosedXML_Sandbox
{
    internal static class Program
    {
        private static void Main(string[] args)
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Table");

                ws.Cell("A1").SetValue("1");
                ws.Cell("A2").SetValue("2");
                ws.Cell("A3").SetValue("3");
                ws.Cell("A4").SetValue("4");
                ws.Cell("A5").SetValue("5");
                ws.Cell("A6").SetValue("6");
                ws.Cell("A7").SetValue("7");
                ws.Cell("A8").SetValue("8");

                ws.Cell("B1").SetValue("1");
                ws.Cell("B2").SetValue("2");
                ws.Cell("B3").SetValue("3");
                ws.Cell("B4").SetValue("4");
                ws.Cell("B5").SetValue("5");
                ws.Cell("B6").SetValue("6");
                ws.Cell("B7").SetValue("7");
                ws.Cell("B8").SetValue("8");

                ws.Cell("C1").FormulaA1 = "=A1+B1";
                ws.Cell("C2").FormulaA1 = "=A2+B2";
                ws.Cell("C3").FormulaA1 = "=A3+B3";
                ws.Cell("C4").FormulaA1 = "=A4+B4";
                ws.Cell("C5").FormulaA1 = "=A5+B5";
                ws.Cell("C6").FormulaA1 = "=A6+B6";
                ws.Cell("C7").FormulaA1 = "=A7+B7";
                ws.Cell("C8").FormulaA1 = "=A8+B8";

                var header = ws.Row(1).InsertRowsAbove(1).First();
                for (Int32 co = 1; co <= ws.LastColumnUsed().ColumnNumber(); co++)
                {
                    header.Cell(co).Value = "Column" + co.ToString();
                }
                var rangeTable = ws.RangeUsed();
                var table = rangeTable.CopyTo(ws.Column(ws.LastColumnUsed().ColumnNumber() + 3)).CreateTable();

                var rangeTable2 = rangeTable.RangeUsed();
                var table2 = rangeTable2.CopyTo(ws.Column(ws.LastColumnUsed().ColumnNumber() + 3)).CreateTable();

                table2.Sort("Column3 Desc");

                wb.SaveAs(@"e:\closedXML\105.xlsx");
            }
            Console.WriteLine("Press any key to continue");
            Console.ReadKey();
        }
    }
}
