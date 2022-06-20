using ClosedXML.Excel;

namespace ClosedXML.Examples.Misc
{
    public class AdjustToContentsWithAutoFilter : IXLExample
    {
        // Public
        public void Create(string filePath)
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("AutoFilter");
            ws.Cell("A1").Value = "AVeryLongColumnHeader";
            ws.Cell("A2").Value = "John";
            ws.Cell("A3").Value = "Hank";
            ws.Cell("A4").Value = "Dagny";

            ws.RangeUsed().SetAutoFilter();

            ws.Columns().AdjustToContents();

            wb.SaveAs(filePath);
        }
    }
}