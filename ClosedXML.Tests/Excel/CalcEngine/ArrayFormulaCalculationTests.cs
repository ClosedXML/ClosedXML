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
                Assert.AreEqual(1, arrayFormulaCell.Value);
            }
        }

        [Test]
        public void SameShapeResultCausesEachCellOfCellGroupToUseCorrespondingValue()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var range = ws.Range("A1:A2");

            range.FormulaArrayA1 = "TRANSPOSE({1,2})";

            Assert.AreEqual(1, ws.Cell("A1").Value);
            Assert.AreEqual(2, ws.Cell("A2").Value);
        }

        [Test]
        public void OnlyLeftmostValuesAreUsedWhenCellGroupHasFewerColumnsThanValue()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var range = ws.Range("A1:C1");

            range.FormulaArrayA1 = "{1,2,3,4,5}";

            Assert.AreEqual(1, ws.Cell("A1").Value);
            Assert.AreEqual(2, ws.Cell("B1").Value);
            Assert.AreEqual(3, ws.Cell("C1").Value);
            Assert.AreEqual(Blank.Value, ws.Cell("D1").Value);
        }

        [Test]
        public void OnlyTopmostValuesAreUsedWhenCellGroupHasFewerRowsThanValue()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var range = ws.Range("A1:A3");

            range.FormulaArrayA1 = "{1;2;3;4;5}";

            Assert.AreEqual(1, ws.Cell("A1").Value);
            Assert.AreEqual(2, ws.Cell("A2").Value);
            Assert.AreEqual(3, ws.Cell("A3").Value);
            Assert.AreEqual(Blank.Value, ws.Cell("A4").Value);
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
                Assert.AreEqual(1, ws.Cell(1, column).Value);
                Assert.AreEqual(2, ws.Cell(2, column).Value);
                Assert.AreEqual(XLError.NoValueAvailable, ws.Cell(3, column).Value);
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
                Assert.AreEqual(1, ws.Cell(row, 1).Value);
                Assert.AreEqual(2, ws.Cell(row, 2).Value);
                Assert.AreEqual(XLError.NoValueAvailable, ws.Cell(row, 3).Value);
            }
        }

        [Test]
        public void ExcessColumnsAndRowsOfCellGroupTakeOnNoValueAvailable()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var range = ws.Range("A1:C3");

            range.FormulaArrayA1 = "{1,2;3,4}";

            Assert.AreEqual(1, ws.Cell("A1").Value);
            Assert.AreEqual(2, ws.Cell("B1").Value);
            Assert.AreEqual(XLError.NoValueAvailable, ws.Cell("C1").Value);
            Assert.AreEqual(3, ws.Cell("A2").Value);
            Assert.AreEqual(4, ws.Cell("B2").Value);
            Assert.AreEqual(XLError.NoValueAvailable, ws.Cell("C2").Value);
            Assert.AreEqual(XLError.NoValueAvailable, ws.Cell("A3").Value);
            Assert.AreEqual(XLError.NoValueAvailable, ws.Cell("B3").Value);
            Assert.AreEqual(XLError.NoValueAvailable, ws.Cell("C3").Value);
        }

        [Test]
        public void Legacy_scalar_functions_evaluate_cells_individually()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = -2;
            ws.Cell("A2").Value = XLError.NameNotRecognized;
            ws.Cell("A3").Value = 2;
            ws.Range("B1:B3").FormulaArrayA1 = "SIGN(A1:A3)";

            Assert.AreEqual(-1, ws.Cell("B1").Value);
            Assert.AreEqual(XLError.NameNotRecognized, ws.Cell("B2").Value);
            Assert.AreEqual(1, ws.Cell("B3").Value);
        }

        [Test]
        public void Legacy_reduce_function_calculate_single_value_from_range_and_broadcasts_the_value_to_array()
        {
            // AVERAGE is a legacy function that takes a range and reduces it to a single value.
            // It calculates a single value and that value is then broadcasts it along the range.
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B1:B4").FormulaArrayA1 = "AVERAGE({-2; #NAME?; 2})";

            Assert.AreEqual(XLError.NameNotRecognized, ws.Cell("B1").Value);
            Assert.AreEqual(XLError.NameNotRecognized, ws.Cell("B2").Value);
            Assert.AreEqual(XLError.NameNotRecognized, ws.Cell("B3").Value);
            Assert.AreEqual(XLError.NameNotRecognized, ws.Cell("B4").Value);
        }
    }
}
