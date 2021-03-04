using ClosedXML.Excel;
using MoreLinq;
using System;
using System.Linq;

namespace ClosedXML.Examples.Sparklines
{
    public class SampleSparklines : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws1 = workbook.AddWorksheet("Linear");

            FillSampleData(ws1);

            ws1.Range("A2:A4").Merge().SetValue("Linear, Colorful 1, All markers, SameForAll scale");
            ws1.SparklineGroups.Add("B2:B4", "C2:P4")
                .SetStyle(XLSparklineTheme.Colorful1)
                .SetShowMarkers(XLSparklineMarkers.All).VerticalAxis
                .SetMaxAxisType(XLSparklineAxisMinMax.SameForAll)
                .SetMinAxisType(XLSparklineAxisMinMax.SameForAll);

            ws1.Range("A5:A7").Merge().SetValue("Linear, Colorful 2, First+Last+High+Low, Automatic scale");
            ws1.SparklineGroups.Add("B5:B7", "C5:P7")
                .SetStyle(XLSparklineTheme.Colorful2)
                .SetShowMarkers(XLSparklineMarkers.FirstPoint | XLSparklineMarkers.LastPoint | XLSparklineMarkers.HighPoint | XLSparklineMarkers.LowPoint);

            ws1.Range("A8:A10").Merge().SetValue("Linear, Colorful 3, Markers+Negative, Custom scale");
            ws1.SparklineGroups.Add("B8:B10", "C8:P10")
                .SetStyle(XLSparklineTheme.Colorful3)
                .SetShowMarkers(XLSparklineMarkers.Markers | XLSparklineMarkers.NegativePoints)
                .VerticalAxis
                .SetManualMax(100)
                .SetManualMin(-80);

            ws1.Range("A11:A13").Merge().SetValue("Linear, Colorful 1, Date range");
            ws1.SparklineGroups.Add("B11:B13", "C11:P13")
                .SetStyle(XLSparklineTheme.Colorful1)
                .SetDateRange(ws1.Range("C1:P1"));

            ws1.Range("A14:A16").Merge().SetValue("Linear, Colorful 4, Line weight=2, Right to left");
            ws1.SparklineGroups.Add("B14:B16", "C14:P16")
                .SetStyle(XLSparklineTheme.Colorful4)
                .SetLineWeight(2)
                .HorizontalAxis
                .SetVisible(true)
                .SetColor(XLColor.Red)
                .SetRightToLeft(true);

            ws1.Range("A17:A19").Merge().SetValue("Linear, Colorful 3, Different ranges");
            ws1.SparklineGroups.Add("B17", "C17:P17")
                .Add("B18", "C18:K18").Single().SparklineGroup
                .Add("B19", "C19:E19").Single().SparklineGroup
                .SetStyle(XLSparklineTheme.Colorful3)
                .SetShowMarkers(XLSparklineMarkers.FirstPoint | XLSparklineMarkers.LastPoint);


            var ws2 = ws1.CopyTo("Column");
            ws2.SparklineGroups.ForEach(g =>
                g.SetType(XLSparklineType.Column));

            ws2.Cell("A2").Value = "Column, Colorful 1, All markers, SameForAll scale";
            ws2.Cell("A5").Value = "Column, Colorful 2, First+Last+High+Low, Automatic scale";
            ws2.Cell("A8").Value = "Column, Colorful 3, Markers+Negative, Custom scale";
            ws2.Cell("A11").Value = "Column, Colorful 1, Date range";
            ws2.Cell("A14").Value = "Column, Colorful 4, Line weight=2, Right to left";
            ws2.Cell("A17").Value = "Column, Colorful 3, Different ranges";


            var ws3 = ws1.CopyTo("Stacked");
            ws3.SparklineGroups.ForEach(g =>
                g.SetType(XLSparklineType.Stacked));

            ws3.Cell("A2").Value = "Stacked, Colorful 1, All markers, SameForAll scale";
            ws3.Cell("A5").Value = "Stacked, Colorful 2, First+Last+High+Low, Automatic scale";
            ws3.Cell("A8").Value = "Stacked, Colorful 3, Markers+Negative, Custom scale";
            ws3.Cell("A11").Value = "Stacked, Colorful 1, Date range";
            ws3.Cell("A14").Value = "Stacked, Colorful 4, Line weight=2, Right to left";
            ws3.Cell("A17").Value = "Stacked, Colorful 3, Different ranges";

            workbook.SaveAs(filePath);
        }

        private void FillSampleData(IXLWorksheet ws)
        {
            ws.Column(1).Style.Alignment.SetWrapText(true);
            ws.Column(1).Width = 30;
            ws.Column(2).Width = 30;

            ws.Range("C1:P1").Cells()
                .ForEach(c => c.Value = new DateTime(2016, 1, 1).AddDays(c.Address.ColumnNumber * 7));

            ws.Range("C2:P19").Cells()
                .ForEach(c => c.Value = Math.Round(c.Address.RowNumber * Math.Sin(c.Address.ColumnNumber) * 10, 0));
        }
    }
}
