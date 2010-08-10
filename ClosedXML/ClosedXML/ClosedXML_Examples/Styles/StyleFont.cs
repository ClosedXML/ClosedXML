using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using System.Drawing;
using ClosedXML.Excel.Style;

namespace ClosedXML_Examples.Styles
{
    public class StyleFont
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Style Font");

            var co = 2;
            var ro = 1;

            
            ws.Cell(++ro, co).Value = "Bold";
            ws.Cell(ro, co).Style.Font.Bold = true;

            ws.Cell(++ro, co).Value = "FontColor - Red";
            ws.Cell(ro, co).Style.Font.FontColor = Color.Red;

            ws.Cell(++ro, co).Value = "FontFamilyNumbering - Script";
            ws.Cell(ro, co).Style.Font.FontFamilyNumbering = XLFontFamilyNumberingValues.Script;

            ws.Cell(++ro, co).Value = "FontName - Arial";
            ws.Cell(ro, co).Style.Font.FontName = "Arial";

            workbook.SaveAs(filePath);
        }
    }
}
