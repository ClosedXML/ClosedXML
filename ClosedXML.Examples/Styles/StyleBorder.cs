using ClosedXML.Excel;

namespace ClosedXML.Examples.Styles
{
    public class StyleBorder : IXLExample
    {
        public void Create(string filePath)
        {
            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Style Border");

            var co = 2;
            var ro = 1;

            ws.Cell(++ro, co).Value = "BottomBorder = Thick; BottomBorderColor = Red";
            ws.Cell(ro, co).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
            ws.Cell(ro, co).Style.Border.BottomBorderColor = XLColor.Red;

            ws.Cell(++ro, co).Value = "LeftBorder = Thick; LeftBorderColor = Blue";
            ws.Cell(ro, co).Style.Border.LeftBorder = XLBorderStyleValues.Thick;
            ws.Cell(ro, co).Style.Border.LeftBorderColor = XLColor.Blue;

            ws.Cell(++ro, co).Value = "TopBorder = Thick; TopBorderColor = Yellow";
            ws.Cell(ro, co).Style.Border.TopBorder = XLBorderStyleValues.Thick;
            ws.Cell(ro, co).Style.Border.TopBorderColor = XLColor.Yellow;

            ws.Cell(++ro, co).Value = "RightBorder = Thick; RightBorderColor = Black";
            ws.Cell(ro, co).Style.Border.RightBorder = XLBorderStyleValues.Thick;
            ws.Cell(ro, co).Style.Border.RightBorderColor = XLColor.Black;

            ws.Cell(++ro, co).Value = "DiagonalBorder = Thin; DiagonalBorderColor = Red; DiagonalUp = true";
            ws.Cell(ro, co).Style.Border.DiagonalBorder = XLBorderStyleValues.Thin;
            ws.Cell(ro, co).Style.Border.DiagonalBorderColor = XLColor.Red;
            ws.Cell(ro, co).Style.Border.DiagonalUp = true;

            ws.Cell(++ro, co).Value = "DiagonalBorder = Thin; DiagonalBorderColor = Red; DiagonalDown = true";
            ws.Cell(ro, co).Style.Border.DiagonalBorder = XLBorderStyleValues.Thin;
            ws.Cell(ro, co).Style.Border.DiagonalBorderColor = XLColor.Red;
            ws.Cell(ro, co).Style.Border.DiagonalDown = true;

            ws.Cell(++ro, co).Value = "DiagonalBorder = Thin; DiagonalBorderColor = Red; DiagonalUp = true; DiagonalDown = true";
            ws.Cell(ro, co).Style.Border.DiagonalBorder = XLBorderStyleValues.Thin;
            ws.Cell(ro, co).Style.Border.DiagonalBorderColor = XLColor.Red;
            ws.Cell(ro, co).Style.Border.DiagonalUp = true;
            ws.Cell(ro, co).Style.Border.DiagonalDown = true;

            workbook.SaveAs(filePath);
        }
    }
}