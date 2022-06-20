using ClosedXML.Excel;

namespace ClosedXML.Examples
{
    public class HelloWorld
    {
        public void Create(string filePath)
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Sample Sheet");
            worksheet.Cell("A1").Value = "Hello World!";
            workbook.SaveAs(filePath);
        }
    }
}