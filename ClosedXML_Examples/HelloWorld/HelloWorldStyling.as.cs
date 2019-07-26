using System;
using ClosedXML.Excel;

namespace ClosedXML_Examples
{
    public class HelloWorld
    {
        public void Create(String filePath)
        {
            IXLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sample Sheet");

            ws.Cell(2,3).Value = "Hello World!";
            ws.Cell(4,2).Value = "Project:";
            ws.Cell(4,4).Value = "ClosedXML Example";
            ws.Cell(6,2).Value = "Author:";
            ws.Cell(6,4).Value = "KnapSac";

            ws.Cell(2,3).Style.Fill.SetBackgroundColor( XLColor.Cyan );

            IXLRange range = ws.Range(ws.Cell(4,2).Address, ws.Cell(6,4).Address);

            range.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;

            wb.SaveAs(filePath);
        }
    }
}
