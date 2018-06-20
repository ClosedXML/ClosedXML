using ClosedXML.Excel;
using System;

namespace ClosedXML_Examples.Sparklines
{
    public class DeletingSparklinesHelp : IXLExample
    {
        // This test does not pass. The example file differs minorly after it is resaved.
        // Only when Clear is called. I can't figure out why.
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();            
            var wsDelete = workbook.AddWorksheet("Delete");

            wsDelete.Cell("A5").Value = "10";
            wsDelete.Cell("B5").Value = "20";
            wsDelete.Cell("C5").Value = "30";
            wsDelete.Cell("D5").Value = "40";
            wsDelete.Cell("E5").Value = "50";

            wsDelete.Cell("A6").Value = "10";
            wsDelete.Cell("B6").Value = "20";
            wsDelete.Cell("C6").Value = "30";
            wsDelete.Cell("D6").Value = "40";
            wsDelete.Cell("E6").Value = "50";

            var slgDelete3 = wsDelete.SparklineGroups.Add(wsDelete);
            slgDelete3.AddSparkline(wsDelete.Cell("F5"), wsDelete.Range("A5:E5").RangeAddress.ToStringRelative(true));
            slgDelete3.AddSparkline(wsDelete.Cell("F6"), wsDelete.Range("A6:E6").RangeAddress.ToStringRelative(true));
            
            wsDelete.Range("F5:F6").Clear();

            workbook.SaveAs(filePath);
        }
    }    
}
