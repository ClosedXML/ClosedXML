using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel.Style;
using ClosedXML.Excel;
using System.Drawing;

namespace ClosedXML_Examples.Styles
{
    public class StyleBorder
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Style Border");

            var co = 2;
            var ro = 1;

            ws.Cell(++ro, co).Value = "BottomBorder = Thick; BottomBorderColor = Red";
            ws.Cell(ro, co).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
            ws.Cell(ro, co).Style.Border.BottomBorderColor = Color.Red;

            ws.Cell(++ro, co).Value = "TopBorder = Thick; TopBorderColor = Red";
            ws.Cell(ro, co).Style.Border.TopBorder = XLBorderStyleValues.Thick;
            ws.Cell(ro, co).Style.Border.TopBorderColor = Color.Red;

            ws.Cell(++ro, co).Value = "LeftBorder = Thick; LeftBorderColor = Red";
            ws.Cell(ro, co).Style.Border.LeftBorder = XLBorderStyleValues.Thick;
            ws.Cell(ro, co).Style.Border.LeftBorderColor = Color.Red;

            ws.Cell(++ro, co).Value = "RightBorder = Thick; RightBorderColor = Red";
            ws.Cell(ro, co).Style.Border.RightBorder = XLBorderStyleValues.Thick;
            ws.Cell(ro, co).Style.Border.RightBorderColor = Color.Red;

            ws.Cell(++ro, co).Value = "DiagonalBorder = Thick; DiagonalBorderColor = Red; DiagonalUp = true";
            ws.Cell(ro, co).Style.Border.DiagonalBorder = XLBorderStyleValues.Thick;
            ws.Cell(ro, co).Style.Border.DiagonalBorderColor = Color.Red;
            ws.Cell(ro, co).Style.Border.DiagonalUp = true;

            ws.Cell(++ro, co).Value = "DiagonalBorder = Thick; DiagonalBorderColor = Red; DiagonalDown = true";
            ws.Cell(ro, co).Style.Border.DiagonalBorder = XLBorderStyleValues.Thick;
            ws.Cell(ro, co).Style.Border.DiagonalBorderColor = Color.Red;
            ws.Cell(ro, co).Style.Border.DiagonalDown = true;

            ws.Cell(++ro, co).Value = "DiagonalBorder = Thick; DiagonalBorderColor = Red; DiagonalUp = true; DiagonalDown = true";
            ws.Cell(ro, co).Style.Border.DiagonalBorder = XLBorderStyleValues.Thick;
            ws.Cell(ro, co).Style.Border.DiagonalBorderColor = Color.Red;
            ws.Cell(ro, co).Style.Border.DiagonalUp = true;
            ws.Cell(ro, co).Style.Border.DiagonalDown = true;

            workbook.SaveAs(filePath);
        }
    }
}