using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    public class CalcEngineCellEnumerationTests
    {
        [Test]
        public void CanEnumerateCellsOverEmptySheet()
        {
            using var wb = new XLWorkbook();
            var sheet1 = wb.AddWorksheet("Sheet1");
            var sheet2 = wb.AddWorksheet("Sheet2");

            var cell = sheet1.FirstCell();
            cell.FormulaA1 = "=SUMIFS(Sheet2!B:B, Sheet2!C:C, 1)";

            Assert.That(cell.Value, Is.EqualTo(0));
        }
    }
}
