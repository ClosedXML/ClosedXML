using ClosedXML.Excel;
using System;

namespace ClosedXML_Examples
{
    public class Sparklines : IXLExample
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

            // Create a worksheet and delete a SparklineGroup from it
            var wsDelete = workbook.AddWorksheet("Delete");

            wsDelete.Cell("A1").Value = "10";
            wsDelete.Cell("B1").Value = "20";
            wsDelete.Cell("C1").Value = "30";
            wsDelete.Cell("D1").Value = "40";
            wsDelete.Cell("E1").Value = "50";

            wsDelete.Cell("A2").Value = "50";
            wsDelete.Cell("B2").Value = "20";
            wsDelete.Cell("C2").Value = "30";
            wsDelete.Cell("D2").Value = "10";
            wsDelete.Cell("E2").Value = "40";

            var slgDelete = wsDelete.SparklineGroups.Add(wsDelete);
            slgDelete.SeriesColor = XLColor.CarrotOrange;
            slgDelete.AddSparkline(wsDelete.Cell("F1"), wsDelete.Range("A1:E1").RangeAddress.ToStringRelative(true));
            slgDelete.AddSparkline(wsDelete.Cell("F2"), wsDelete.Range("A2:E2").RangeAddress.ToStringRelative(true));

            wsDelete.SparklineGroups.Remove(slgDelete);
            
            // Create a worksheet add a sparkline to a cell, remove the cell
            var wsRemoveCell = workbook.AddWorksheet("RemoveCell");

            wsRemoveCell.Cell("A1").Value = "10";
            wsRemoveCell.Cell("B1").Value = "20";
            wsRemoveCell.Cell("C1").Value = "30";
            wsRemoveCell.Cell("D1").Value = "40";
            wsRemoveCell.Cell("E1").Value = "50";

            var slgRemoveCell = wsRemoveCell.SparklineGroups.Add(wsRemoveCell);
            slgRemoveCell.SeriesColor = XLColor.CarrotOrange;
            slgRemoveCell.AddSparkline(wsRemoveCell.Cell("F1"), wsRemoveCell.Range("A1:E1").RangeAddress.ToStringRelative(true));

            wsRemoveCell.Cell("F1").Delete(XLShiftDeletedCells.ShiftCellsLeft);

            workbook.SaveAs(filePath);
        }
    }    
}
