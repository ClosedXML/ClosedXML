using System;
using ClosedXML.Excel;

namespace ClosedXML_Examples
{
    public class HelloWorld
    {
        public void Create(String filePath)
        {
            IXLWorkbook workbook = new XLWorkbook();
            IXLWorksheet worksheet = workbook.Worksheets.Add("Sample Sheet");
            worksheet.Cell(1,1).Value = "Hello World!";
            workbook.SaveAs(filePath);
        }
    }
}
