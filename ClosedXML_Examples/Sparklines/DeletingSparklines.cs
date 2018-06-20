using ClosedXML.Excel;
using System;

namespace ClosedXML_Examples.Sparklines
{
    public class DeletingSparklines : IXLExample
    {
        // Create a worksheet add sparklines and delete them
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();            
            var wsDelete = workbook.AddWorksheet("Delete");

            wsDelete.Cell("A1").Value = "10";
            wsDelete.Cell("B1").Value = "20";
            wsDelete.Cell("C1").Value = "30";
            wsDelete.Cell("D1").Value = "40";
            wsDelete.Cell("E1").Value = "50";

            wsDelete.Cell("A2").Value = "10";
            wsDelete.Cell("B2").Value = "20";
            wsDelete.Cell("C2").Value = "30";
            wsDelete.Cell("D2").Value = "40";
            wsDelete.Cell("E2").Value = "50";

            wsDelete.Cell("A3").Value = "10";
            wsDelete.Cell("B3").Value = "20";
            wsDelete.Cell("C3").Value = "30";
            wsDelete.Cell("D3").Value = "40";
            wsDelete.Cell("E3").Value = "50";

            wsDelete.Cell("A4").Value = "10";
            wsDelete.Cell("B4").Value = "20";
            wsDelete.Cell("C4").Value = "30";
            wsDelete.Cell("D4").Value = "40";
            wsDelete.Cell("E4").Value = "50";

            //wsDelete.Cell("A5").Value = "10";
            //wsDelete.Cell("B5").Value = "20";
            //wsDelete.Cell("C5").Value = "30";
            //wsDelete.Cell("D5").Value = "40";
            //wsDelete.Cell("E5").Value = "50";

            //wsDelete.Cell("A6").Value = "10";
            //wsDelete.Cell("B6").Value = "20";
            //wsDelete.Cell("C6").Value = "30";
            //wsDelete.Cell("D6").Value = "40";
            //wsDelete.Cell("E6").Value = "50";

            var slgDelete1 = wsDelete.SparklineGroups.Add(wsDelete);
            slgDelete1.AddSparkline(wsDelete.Cell("F1"), wsDelete.Range("A1:E1").RangeAddress.ToStringRelative(true));
            slgDelete1.AddSparkline(wsDelete.Cell("F2"), wsDelete.Range("A2:E2").RangeAddress.ToStringRelative(true));

            var slgDelete2 = wsDelete.SparklineGroups.Add(wsDelete);
            var slDelete = slgDelete2.AddSparkline(wsDelete.Cell("F3"), wsDelete.Range("A3:E3").RangeAddress.ToStringRelative(true));
            slgDelete2.AddSparkline(wsDelete.Cell("F4"), wsDelete.Range("A4:E4").RangeAddress.ToStringRelative(true));

            var slgDelete3 = wsDelete.SparklineGroups.Add(wsDelete);
            slgDelete3.AddSparkline(wsDelete.Cell("F5"), wsDelete.Range("A5:E5").RangeAddress.ToStringRelative(true));
            slgDelete3.AddSparkline(wsDelete.Cell("F6"), wsDelete.Range("A6:E6").RangeAddress.ToStringRelative(true));
            
            wsDelete.SparklineGroups.Remove(slgDelete1);
            wsDelete.SparklineGroups.Remove(slDelete);
            wsDelete.SparklineGroups.Remove(wsDelete.Cell("F4"));
            //wsDelete.Range("F5:F6").Clear();

            workbook.SaveAs(filePath);
        }
    }    
}
