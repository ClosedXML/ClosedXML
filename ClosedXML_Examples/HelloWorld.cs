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

            ws.Cell(1,1).Value = "Hello World!";

            wb.SaveAs(filePath);
        }
    }
}
