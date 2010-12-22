using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using System.Drawing;


namespace ClosedXML_Examples.Styles
{
    public class UsingColors
    {
        public void Create(String filePath)
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Using Colors");

            // From Known color
            ws.Cell(1, 1).Style.Fill.BackgroundColor = XLColor.Red;
            ws.Cell(1, 2).Value = "XLColor.Red";

            // From Color not known
            ws.Cell(2, 1).Style.Fill.BackgroundColor = XLColor.Byzantine;
            ws.Cell(2, 2).Value = "XLColor.Byzantine";

            // From Theme color
            ws.Cell(3, 1).Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent1);
            ws.Cell(3, 2).Value = "XLColor.FromTheme(XLThemeColor.Accent1)";

            // From Theme color with tint
            ws.Cell(4, 1).Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent2, 0.5);
            ws.Cell(4, 2).Value = "XLColor.FromTheme(XLThemeColor.Accent2, 0.5)";

            // From indexed color (legacy)
            ws.Cell(5, 1).Style.Fill.BackgroundColor = XLColor.FromIndex(25);
            ws.Cell(5, 2).Value = "XLColor.FromIndex(25)";

            ws.Columns().AdjustToContents();

            wb.SaveAs(filePath);
        }
    }
}