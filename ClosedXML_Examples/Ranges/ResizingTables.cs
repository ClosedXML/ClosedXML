using ClosedXML.Excel;
using System.IO;
using System.Linq;

namespace ClosedXML_Examples.Ranges
{
    public class ResizingTables : IXLExample
    {
        public void Create(string filePath)
        {
            string tempFile = ExampleHelper.GetTempFilePath(filePath);
            try
            {
                new UsingTables().Create(tempFile);
                using (var wb = new XLWorkbook(tempFile))
                {
                    var ws1 = wb.Worksheets.First();

                    var ws2 = ws1.CopyTo("Contacts 2");
                    ws2.Cell("A2").Value = "Index";
                    ws2.Cell("A3").Value = Enumerable.Range(1, 3).ToArray();
                    var table2 = ws2.Tables.First().SetShowTotalsRow(false);
                    table2.Resize(ws2.Range(ws2.Cell("A2"), table2.DataRange.LastCell()));

                    var ws3 = ws1.CopyTo("Contacts 3");
                    var table3 = ws3.Tables.First().SetShowTotalsRow(false);
                    table3.Resize(ws3.Range(table3.AsRange().FirstCell(), table3.DataRange.LastCell().CellLeft()));

                    wb.SaveAs(filePath);
                }
            }
            finally
            {
                if (File.Exists(tempFile))
                {
                    File.Delete(tempFile);
                }
            }
        }
    }
}
