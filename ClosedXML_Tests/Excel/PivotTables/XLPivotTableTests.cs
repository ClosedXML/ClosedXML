using ClosedXML.Excel;
using NUnit.Framework;
using System.IO;

namespace ClosedXML_Tests
{
    [TestFixture]
    public class XLPivotTableTests
    {
        [Test]
        public void PivotTables()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\PivotTables\PivotTables.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheet("PastrySalesData");
                var table = ws.Table("PastrySalesData");

                var range = table.DataRange;
                var header = ws.Range(1, 1, 1, 3);
                var dataRange = ws.Range(header.FirstCell(), range.LastCell());

                var ptSheet = wb.Worksheets.Add("BlankPivotTable");
                var pt = ptSheet.PivotTables.AddNew("pvt", ptSheet.Cell(1, 1), dataRange);

                using (var ms = new MemoryStream())
                {
                    wb.SaveAs(ms, true);
                }
            }
        }
    }
}
