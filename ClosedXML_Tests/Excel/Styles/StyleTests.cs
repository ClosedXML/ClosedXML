using ClosedXML.Excel;
using NUnit.Framework;
using System.IO;
using System.Linq;

namespace ClosedXML_Tests.Excel
{
    [TestFixture]
    public class StyleTests
    {
        [Test]
        public void EmptyCellWithQuotePrefixNotTreatedAsEmpty()
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = new XLWorkbook())
                {
                    var ws = wb.AddWorksheet("Sheet1");
                    ws.FirstCell().SetValue("Empty cell with quote prefix:");
                    var cell = ws.FirstCell().CellRight() as XLCell;

                    Assert.IsTrue(cell.IsEmpty());
                    cell.Style.IncludeQuotePrefix = true;

                    Assert.IsTrue(cell.IsEmpty());
                    Assert.IsFalse(cell.IsEmpty(XLCellsUsedOptions.All));

                    wb.SaveAs(ms);
                }

                ms.Seek(0, SeekOrigin.Begin);

                using (var wb = new XLWorkbook(ms))
                {
                    var ws = wb.Worksheets.First();
                    var cell = ws.FirstCell().CellRight() as XLCell;
                    Assert.AreEqual(1, cell.SharedStringId);

                    Assert.IsTrue(cell.IsEmpty());
                    Assert.IsFalse(cell.IsEmpty(XLCellsUsedOptions.All));
                }
            }
        }

        [TestCase("A1", TestName = "First cell")]
        [TestCase("A2", TestName = "Cell from initialized row")]
        [TestCase("B1", TestName = "Cell from initialized column")]
        [TestCase("D4", TestName = "Initialized cell")]
        [TestCase("F6", TestName = "Non-initialized cell")]
        public void CellTakesWorksheetStyle(string cellAddress)
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.Column(2);
                ws.Row(2);
                ws.Cell("D4").Value = "Non empty";
                ws.Style.Font.SetFontName("Arial");
                ws.Style.Font.SetFontSize(9);

                var cell = ws.Cell(cellAddress);
                Assert.AreEqual("Arial", cell.Style.Font.FontName);
                Assert.AreEqual(9, cell.Style.Font.FontSize);
            }
        }
    }
}
