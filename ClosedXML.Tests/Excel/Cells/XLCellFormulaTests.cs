using NUnit.Framework;
using ClosedXML.Excel;

namespace ClosedXML.Tests.Excel.Cells
{
    [TestFixture]
    public class XLCellFormulaTests
    {
        [Test]
        public void CellFormulaIsStrippedOfEqualSign()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell(1, 1).FormulaA1 = "=B1";
            Assert.AreEqual("B1", ws.Cell(1, 1).FormulaA1);
        }

        [Test]
        public void DataTable_MaintainProperties()
        {
            TestHelper.LoadSaveAndCompare(
                @"Other\Formulas\DataTableFormula-Excel-Input.xlsx",
                @"Other\Formulas\DataTableFormula-Output.xlsx");
        }
    }
}
