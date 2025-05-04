using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    /// <summary>
    /// Tests that calc engine adjusts its internal state in response to changes of workbook structure.
    /// </summary>
    [TestFixture]
    internal class CalcEngineListenerTests
    {
        [Test]
        public void Formulas_dependent_on_specific_sheet_are_dirty_after_sheet_addition()
        {
            using var wb = new XLWorkbook();
            var sutWs = wb.AddWorksheet();
            sutWs.Cell("A1").FormulaA1 = "new!A1";
            Assert.That(sutWs.Cell("A1").Value, Is.EqualTo(XLError.CellReference));

            var newWs = wb.AddWorksheet("new");
            newWs.Cell("A1").Value = 5;

            Assert.Multiple(() =>
            {
                // Cell contains last calculated value
                Assert.That(sutWs.Cell("A1").CachedValue, Is.EqualTo(XLError.CellReference));

                // But once asked for real value, it calculates it.
                Assert.That(sutWs.Cell("A1").NeedsRecalculation, Is.True);
                Assert.That(sutWs.Cell("A1").Value, Is.EqualTo(5.0));
            });
        }

        [Test]
        public void Formulas_dependent_on_specific_sheet_are_dirty_after_sheet_deletion()
        {
            using var wb = new XLWorkbook();
            var keptWs = wb.AddWorksheet();
            var deletedWs = wb.AddWorksheet("deleted");

            deletedWs.Cell("A1").Value = 5;
            keptWs.Cell("A1").FormulaA1 = "deleted!A1";
            Assert.That(keptWs.Cell("A1").Value, Is.EqualTo(5.0));

            deletedWs.Delete();

            Assert.Multiple(() =>
            {
                // Cell contains last calculated value
                Assert.That(keptWs.Cell("A1").CachedValue, Is.EqualTo(5.0));

                // But once asked for real value, it calculates it.
                Assert.That(keptWs.Cell("A1").NeedsRecalculation, Is.True);
                Assert.That(keptWs.Cell("A1").Value, Is.EqualTo(XLError.CellReference));
            });
        }

        [Test]
        public void Formulas_are_shifted_when_area_is_added_and_cells_shifted_down()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").FormulaA1 = "B1*2";
            ws.Cell("B1").FormulaA1 = "C1*2";
            ws.Cell("C1").FormulaA1 = "1+2";

            ws.RecalculateAllFormulas();

            ws.Range("A1:B1").InsertRowsAbove(2);

            Assert.That(ws.Cell("A3").Value, Is.EqualTo(12.0));
            Assert.That(ws.Cell("A3").NeedsRecalculation, Is.False);
            Assert.That(ws.Cell("B3").NeedsRecalculation, Is.False);

            // Dependency tree should pick up the change
            ws.Cell("C1").FormulaA1 = "2+2";
            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell("A3").NeedsRecalculation, Is.True);
                Assert.That(ws.Cell("B3").NeedsRecalculation, Is.True);
                Assert.That(ws.Cell("A3").Value, Is.EqualTo(16.0));
            });
        }

        [Test]
        public void Formulas_are_shifted_when_area_is_added_and_cells_shifted_right()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").FormulaA1 = "A2*2";
            ws.Cell("A2").FormulaA1 = "A3*2";
            ws.Cell("A3").FormulaA1 = "1+2";

            ws.RecalculateAllFormulas();

            ws.Cell("A2").InsertCellsBefore(4);

            Assert.That(ws.Cell("A1").Value, Is.EqualTo(12.0));
            Assert.That(ws.Cell("E2").NeedsRecalculation, Is.False);

            // Dependency tree should pick up the change
            ws.Cell("A3").FormulaA1 = "2+2";
            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell("E2").NeedsRecalculation, Is.True);
                Assert.That(ws.Cell("A1").NeedsRecalculation, Is.True);
                Assert.That(ws.Cell("A1").Value, Is.EqualTo(16.0));
            });
        }

        [Test]
        public void Formulas_are_shifted_when_area_is_deleted_and_cells_shifted_up()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A5").FormulaA1 = "1+2";
            ws.Cell("B5").FormulaA1 = "A5*2";
            ws.Cell("C5").FormulaA1 = "B5*2";

            ws.RecalculateAllFormulas();

            ws.Range("B2:C4").Delete(XLShiftDeletedCells.ShiftCellsUp);

            Assert.That(ws.Cell("C2").Value, Is.EqualTo(12.0));
            Assert.That(ws.Cell("B2").NeedsRecalculation, Is.False);
            Assert.That(ws.Cell("A2").NeedsRecalculation, Is.False);

            // Dependency tree should pick up the change
            ws.Cell("A5").FormulaA1 = "2+2";
            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell("B2").NeedsRecalculation, Is.True);
                Assert.That(ws.Cell("C2").NeedsRecalculation, Is.True);
                Assert.That(ws.Cell("C2").Value, Is.EqualTo(16.0));
            });
        }

        [Test]
        public void Formulas_are_shifted_when_area_is_deleted_and_cells_shifted_left()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("D1").FormulaA1 = "1+2";
            ws.Cell("E2").FormulaA1 = "D1*2";
            ws.Cell("D3").FormulaA1 = "E2*2";

            ws.RecalculateAllFormulas();

            ws.Range("A1:C5").Delete(XLShiftDeletedCells.ShiftCellsLeft);

            Assert.That(ws.Cell("A3").Value, Is.EqualTo(12.0));
            Assert.That(ws.Cell("B2").NeedsRecalculation, Is.False);
            Assert.That(ws.Cell("A1").NeedsRecalculation, Is.False);

            // Dependency tree should pick up the change
            ws.Cell("A1").FormulaA1 = "2+2";
            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell("B2").NeedsRecalculation, Is.True);
                Assert.That(ws.Cell("A3").NeedsRecalculation, Is.True);
                Assert.That(ws.Cell("A3").Value, Is.EqualTo(16.0));
            });
        }
    }
}
