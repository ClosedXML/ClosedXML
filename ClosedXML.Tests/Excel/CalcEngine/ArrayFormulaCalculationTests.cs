using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    public class ArrayFormulaCalculationTests
    {
        [Test]
        public void ScalarResultOfArrayFormulaIsCopiedAcrossCellGroup()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var range = ws.Range("C2:D4");

            range.FormulaArrayA1 = "ABS(-1)";

            foreach (var arrayFormulaCell in range.Cells())
            {
                Assert.That(arrayFormulaCell.Value, Is.EqualTo(1));
            }
        }

        [Test]
        public void SameShapeResultCausesEachCellOfCellGroupToUseCorrespondingValue()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var range = ws.Range("A1:A2");

            range.FormulaArrayA1 = "TRANSPOSE({1,2})";

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell("A1").Value, Is.EqualTo(1));
                Assert.That(ws.Cell("A2").Value, Is.EqualTo(2));
            });
        }

        [Test]
        public void OnlyLeftmostValuesAreUsedWhenCellGroupHasFewerColumnsThanValue()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var range = ws.Range("A1:C1");

            range.FormulaArrayA1 = "{1,2,3,4,5}";

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell("A1").Value, Is.EqualTo(1));
                Assert.That(ws.Cell("B1").Value, Is.EqualTo(2));
                Assert.That(ws.Cell("C1").Value, Is.EqualTo(3));
                Assert.That(ws.Cell("D1").Value, Is.EqualTo(Blank.Value));
            });
        }

        [Test]
        public void OnlyTopmostValuesAreUsedWhenCellGroupHasFewerRowsThanValue()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var range = ws.Range("A1:A3");

            range.FormulaArrayA1 = "{1;2;3;4;5}";

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell("A1").Value, Is.EqualTo(1));
                Assert.That(ws.Cell("A2").Value, Is.EqualTo(2));
                Assert.That(ws.Cell("A3").Value, Is.EqualTo(3));
                Assert.That(ws.Cell("A4").Value, Is.EqualTo(Blank.Value));
            });
        }

        [Test]
        public void SingleColumnValueIsClonedAcrossCellGroup()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var range = ws.Range("A1:C3");

            range.FormulaArrayA1 = "{1;2}";

            for (var column = 1; column <= 3; column++)
            {
                Assert.Multiple(() =>
                {
                    Assert.That(ws.Cell(1, column).Value, Is.EqualTo(1));
                    Assert.That(ws.Cell(2, column).Value, Is.EqualTo(2));
                    Assert.That(ws.Cell(3, column).Value, Is.EqualTo(XLError.NoValueAvailable));
                });
            }
        }

        [Test]
        public void SingleRowValueIsClonedAcrossCellGroup()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var range = ws.Range("A1:C3");

            range.FormulaArrayA1 = "{1,2}";

            for (var row = 1; row <= 3; row++)
            {
                Assert.Multiple(() =>
                {
                    Assert.That(ws.Cell(row, 1).Value, Is.EqualTo(1));
                    Assert.That(ws.Cell(row, 2).Value, Is.EqualTo(2));
                    Assert.That(ws.Cell(row, 3).Value, Is.EqualTo(XLError.NoValueAvailable));
                });
            }
        }

        [Test]
        public void ExcessColumnsAndRowsOfCellGroupTakeOnNoValueAvailable()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var range = ws.Range("A1:C3");

            range.FormulaArrayA1 = "{1,2;3,4}";

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell("A1").Value, Is.EqualTo(1));
                Assert.That(ws.Cell("B1").Value, Is.EqualTo(2));
                Assert.That(ws.Cell("C1").Value, Is.EqualTo(XLError.NoValueAvailable));
                Assert.That(ws.Cell("A2").Value, Is.EqualTo(3));
                Assert.That(ws.Cell("B2").Value, Is.EqualTo(4));
                Assert.That(ws.Cell("C2").Value, Is.EqualTo(XLError.NoValueAvailable));
                Assert.That(ws.Cell("A3").Value, Is.EqualTo(XLError.NoValueAvailable));
                Assert.That(ws.Cell("B3").Value, Is.EqualTo(XLError.NoValueAvailable));
                Assert.That(ws.Cell("C3").Value, Is.EqualTo(XLError.NoValueAvailable));
            });
        }

        [Test]
        public void Array_argument_for_scalar_function_in_array_formula_uses_only_first_value_of_array()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B1:B3").FormulaArrayA1 = "SIGN({-1,2,0})";

            Assert.Multiple(() =>
            {
                // Uses only -1 for all values
                Assert.That(ws.Cell("B1").Value, Is.EqualTo(-1));
                Assert.That(ws.Cell("B2").Value, Is.EqualTo(-1));
                Assert.That(ws.Cell("B3").Value, Is.EqualTo(-1));
            });
        }
    }
}
