using ClosedXML.Excel;

namespace ClosedXML.Examples.Styles
{
    public class StyleFont : IXLExample
    {
        public void Create(string filePath)
        {
            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Style Font");

            var co = 2;
            var ro = 1;

            ws.Cell(++ro, co).Value = "Bold";
            ws.Cell(ro, co).Style.Font.Bold = true;

            ws.Cell(++ro, co).Value = "FontColor - Red";
            ws.Cell(ro, co).Style.Font.FontColor = XLColor.Red;

            ws.Cell(++ro, co).Value = "FontFamilyNumbering - Script";
            ws.Cell(ro, co).Style.Font.FontFamilyNumbering = XLFontFamilyNumberingValues.Script;

            ws.Cell(++ro, co).Value = "FontCharSet - العربية التنضيد";
            ws.Cell(ro, co).Style
                .Font.SetFontName("Arabic Typesetting")
                .Font.SetFontCharSet(XLFontCharSet.Arabic);

            ws.Cell(++ro, co).Value = "FontName - Stencil";
            ws.Cell(ro, co).Style.Font.FontName = "Stencil";

            ws.Cell(++ro, co).Value = "FontSize - 15";
            ws.Cell(ro, co).Style.Font.FontSize = 15;

            ws.Cell(++ro, co).Value = "Italic - true";
            ws.Cell(ro, co).Style.Font.Italic = true;

            ws.Cell(++ro, co).Value = "Strikethrough - true";
            ws.Cell(ro, co).Style.Font.Strikethrough = true;

            ws.Cell(++ro, co).Value = "Underline - Double";
            ws.Cell(ro, co).Style.Font.Underline = XLFontUnderlineValues.Double;

            ws.Cell(++ro, co).Value = "VerticalAlignment - Superscript";
            ws.Cell(ro, co).Style.Font.VerticalAlignment = XLFontVerticalTextAlignmentValues.Superscript;

            ws.Column(co).AdjustToContents();

            workbook.SaveAs(filePath);
        }
    }
}