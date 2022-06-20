using ClosedXML.Excel;

namespace ClosedXML.Examples.Styles
{
    public class StyleFill : IXLExample
    {
        public void Create(string filePath)
        {
            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Style Fill");

            var co = 2;
            var ro = 1;

            ws.Cell(++ro, co + 1).Value = "BackgroundColor = Red";
            ws.Cell(ro, co).Style.Fill.BackgroundColor = XLColor.Red;

            ws.Cell(++ro, co + 1).Value = "PatternType = DarkTrellis; PatternColor = Orange; BackgroundColor = Blue";
            ws.Cell(ro, co).Style.Fill.PatternType = XLFillPatternValues.DarkTrellis;
            ws.Cell(ro, co).Style.Fill.PatternColor = XLColor.Orange;
            ws.Cell(ro, co).Style.Fill.BackgroundColor = XLColor.Blue;

            workbook.SaveAs(filePath);
        }
    }
}