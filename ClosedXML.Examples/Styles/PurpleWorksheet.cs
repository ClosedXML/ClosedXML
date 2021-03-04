using System;
using ClosedXML.Excel;


namespace ClosedXML.Examples.Styles
{
    public class PurpleWorksheet : IXLExample
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Purple Worksheet");

            ws.Style.Fill.BackgroundColor = XLColor.Purple;
            
            workbook.SaveAs(filePath);
        }
    }
}