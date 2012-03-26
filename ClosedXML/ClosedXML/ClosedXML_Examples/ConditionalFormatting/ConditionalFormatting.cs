using System;
using ClosedXML.Excel;


namespace ClosedXML_Examples
{
    public class CFColorScaleLowMidHigh : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue(1)
                .CellBelow().SetValue(1)
                .CellBelow().SetValue(2)
                .CellBelow().SetValue(3);

            ws.RangeUsed().AddConditionalFormat().ColorScale()
                .LowestValue(XLColor.Red)
                .Midpoint(XLCFContentType.Percent, "50", XLColor.Yellow) 
                .HighestValue(XLColor.Green);

            workbook.SaveAs(filePath);
        }
    }

    public class CFColorScaleLowHigh : IXLExample
    {

        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue(1)
                .CellBelow().SetValue(1)
                .CellBelow().SetValue(2)
                .CellBelow().SetValue(3);

            ws.RangeUsed().AddConditionalFormat().ColorScale()
                .Minimum(XLCFContentType.Number, "2", XLColor.Red)
                .Maximum(XLCFContentType.Percentile, "90", XLColor.Green);

            workbook.SaveAs(filePath);
        }
    }

    public class CFStartsWith : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet("Sheet1");

            ws.FirstCell().SetValue("Hello")
                .CellBelow().SetValue("Hellos")
                .CellBelow().SetValue("Hell")
                .CellBelow().SetValue("Holl");

            ws.RangeUsed().AddConditionalFormat().WhenStartsWith("Hell")
                .Fill.SetBackgroundColor(XLColor.Red)
                .Border.SetOutsideBorder(XLBorderStyleValues.Thick)
                .Border.SetOutsideBorderColor(XLColor.Blue);

            workbook.SaveAs(filePath);
        }
    }
}
