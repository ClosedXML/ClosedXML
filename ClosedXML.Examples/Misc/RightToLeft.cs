using ClosedXML.Excel;

namespace ClosedXML.Examples.Misc
{
    public class RightToLeft : IXLExample
    {
        public void Create(string filePath)
        {
            using var wb = new XLWorkbook();

            var ws = wb.Worksheets.Add("RightToLeftSheet");
            ws.Cell("A1").Value = "A1";
            ws.Cell("B1").Value = "B1";
            ws.Cell("C1").Value = "C1";
            ws.RightToLeft = true;

            wb.SaveAs(filePath);
        }
    }
}