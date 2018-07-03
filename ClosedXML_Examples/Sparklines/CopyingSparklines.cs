using ClosedXML.Excel;
using System;

namespace ClosedXML_Examples.Sparklines
{
    public class CopyingSparklines : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();

            var ws = workbook.AddWorksheet("Sparklines");

            ws.Cell("A1").Value = "10";
            ws.Cell("B1").Value = "20";
            ws.Cell("C1").Value = "30";
            ws.Cell("D1").Value = "40";
            ws.Cell("E1").Value = "50";

            ws.Cell("A2").Value = "50";
            ws.Cell("B2").Value = "20";
            ws.Cell("C2").Value = "30";
            ws.Cell("D2").Value = "10";
            ws.Cell("E2").Value = "40";

            var slg = ws.SparklineGroups.Add(ws.Range("F1:F2"), ws.Range("A1:E2"));
            slg.SeriesColor = XLColor.CarrotOrange;

            ws.Cell("A4").Value = "10";
            ws.Cell("B4").Value = "20";
            ws.Cell("C4").Value = "30";
            ws.Cell("D4").Value = "40";
            ws.Cell("E4").Value = "50";

            ws.Cell("A5").Value = "50";
            ws.Cell("B5").Value = "20";
            ws.Cell("C5").Value = "30";
            ws.Cell("D5").Value = "10";
            ws.Cell("E5").Value = "40";

            var slg2 = ws.SparklineGroups.Add(ws.Range("F4:F5"), ws.Range("A4:E5"));
            slg2.Type = XLSparklineType.Column;
            slg2.SeriesColor = XLColor.Red;

            ws.Cell("A7").Value = "1";
            ws.Cell("B7").Value = "-1";
            ws.Cell("C7").Value = "-1";
            ws.Cell("D7").Value = "1";
            ws.Cell("E7").Value = "-1";

            ws.Cell("A8").Value = "1";
            ws.Cell("B8").Value = "-1";
            ws.Cell("C8").Value = "-1";
            ws.Cell("D8").Value = "1";
            ws.Cell("E8").Value = "-1";

            var slg3 = ws.SparklineGroups.Add(ws.Range("F7:F8"), ws.Range("A7:E8"));
            slg3.Type = XLSparklineType.Stacked;
            slg3.SeriesColor = XLColor.VividViolet;
            
            //Copy worksheet and ensure the sparkline groups are copied
            ws.CopyTo("CopyWorksheet");

            //Create a worksheet and copy a Sparkline group to it
            var wsCopySparklineGroup = workbook.AddWorksheet("CopySparklineGroup");

            wsCopySparklineGroup.Cell("A1").Value = "10";
            wsCopySparklineGroup.Cell("B1").Value = "20";
            wsCopySparklineGroup.Cell("C1").Value = "30";
            wsCopySparklineGroup.Cell("D1").Value = "40";
            wsCopySparklineGroup.Cell("E1").Value = "50";

            wsCopySparklineGroup.Cell("A2").Value = "50";
            wsCopySparklineGroup.Cell("B2").Value = "20";
            wsCopySparklineGroup.Cell("C2").Value = "30";
            wsCopySparklineGroup.Cell("D2").Value = "10";
            wsCopySparklineGroup.Cell("E2").Value = "40";
            
            slg.CopyTo(wsCopySparklineGroup);

            //Create a worksheet and copy a range with sparklines to it
            var wsCopyRangeWithSparklines = workbook.AddWorksheet("CopyRangeWithSparklines");

            ws.Range("A1:F8").CopyTo(wsCopyRangeWithSparklines.Range("A1:F8"));

            workbook.SaveAs(filePath);
        }
    }    
}
