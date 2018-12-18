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
    }
}
