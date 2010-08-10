using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using System.Drawing;
using ClosedXML.Excel.Style;

namespace ClosedXML_Examples.Styles
{
    public class StyleFill
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Style Fill");

            var co = 2;
            var ro = 1;

            ws.Cell(++ro, co).Value = "BackgroundColor = Red";
            ws.Cell(ro, co).Style.Fill.BackgroundColor = Color.Red;

            ws.Cell(++ro, co).Value = "PatternType = DarkTrellis; PatternColor = Red; PatternBackgroundColor = Blue";
            ws.Cell(ro, co).Style.Fill.PatternType = XLFillPatternValues.DarkTrellis;
            ws.Cell(ro, co).Style.Fill.PatternColor = Color.Red;
            ws.Cell(ro, co).Style.Fill.PatternBackgroundColor = Color.Blue;

            workbook.SaveAs(filePath);
        }
    }
}
