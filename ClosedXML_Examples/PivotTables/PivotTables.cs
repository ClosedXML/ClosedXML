using System;
using ClosedXML.Excel;

namespace ClosedXML_Examples
{
    public class PivotTables
    {
        public void Create(String filePath)
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Pivot Table");
            
            ws.Cell("A1").Value = "Category";
            ws.Cell("A2").Value = "A";
            ws.Cell("A3").Value = "B";
            ws.Cell("A4").Value = "B";
            ws.Cell("B1").Value = "Number";
            ws.Cell("B2").Value = 100;
            ws.Cell("B3").Value = 150;
            ws.Cell("B4").Value = 75;

            //var pivotTable = ws.Range("A1:B4").CreatePivotTable(ws.Cell("D1"));
            //pivotTable.RowLabels.Add("Category");
            //pivotTable.Values.Add("Number")
            //    .ShowAsPctFrom("Category").And("A")
            //    .NumberFormat.Format = "0%";

            wb.SaveAs(filePath);
        }
    }
}
