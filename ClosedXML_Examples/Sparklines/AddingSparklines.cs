using ClosedXML.Excel;
using System;

namespace ClosedXML_Examples.Sparklines
{
    public class AddingSparklines : IXLExample
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

            var slg = ws.SparklineGroups.Add(ws);
            slg.SeriesColor = XLColor.CarrotOrange;
            slg.AddSparkline(ws.Cell("F1"), ws.Range("A1:E1").RangeAddress.ToStringRelative(true));
            slg.AddSparkline(ws.Cell("F2"), ws.Range("A2:E2").RangeAddress.ToStringRelative(true));

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

            var slg2 = ws.SparklineGroups.Add(ws);
            slg2.Type = XLSparklineType.Column;
            slg2.SeriesColor = XLColor.Red;
            slg2.AddSparkline(ws.Cell("F4"), ws.Range("A4:E4").RangeAddress.ToStringRelative(true));
            slg2.AddSparkline(ws.Cell("F5"), ws.Range("A5:E5").RangeAddress.ToStringRelative(true));

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

            var slg3 = ws.SparklineGroups.Add(ws);
            slg3.Type = XLSparklineType.Stacked;
            slg3.SeriesColor = XLColor.VividViolet;
            slg3.AddSparkline(ws.Cell("F7"), ws.Range("A7:E7").RangeAddress.ToStringRelative(true));
            slg3.AddSparkline(ws.Cell("F8"), ws.Range("A8:E8").RangeAddress.ToStringRelative(true));            

            workbook.SaveAs(filePath);
        }
    }    
}
