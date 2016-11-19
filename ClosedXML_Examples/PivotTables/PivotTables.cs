using ClosedXML.Excel;
using System;

namespace ClosedXML_Examples
{
    public class PivotTables : IXLExample
    {
        public void Create(String filePath)
        {
            var wb = new XLWorkbook();

            var wsData = wb.Worksheets.Add("Data");
            wsData.Cell("A1").Value = "Category";
            wsData.Cell("A2").Value = "A";
            wsData.Cell("A3").Value = "B";
            wsData.Cell("A4").Value = "B";
            wsData.Cell("B1").Value = "Number";
            wsData.Cell("B2").Value = 100;
            wsData.Cell("B3").Value = 150;
            wsData.Cell("B4").Value = 75;
            var source = wsData.Range("A1:B4");

            for (int i = 1; i <= 3; i++)
            {
                var name = "PT" + i;
                var wsPT = wb.Worksheets.Add(name);
                var pt = wsPT.PivotTables.AddNew(name, wsPT.Cell("A1"), source);
                pt.RowLabels.Add("Category");
                pt.Values.Add("Number")
                    .ShowAsPctFrom("Category").And("A")
                    .NumberFormat.Format = "0%";
            }

            wb.SaveAs(filePath);
        }
    }
}
