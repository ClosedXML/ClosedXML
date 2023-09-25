using System;
using System.Linq;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    public class ArrayFormulaTests
    {
        [Test]
        public void ArrayFormulaIsSaved()
        {
            TestHelper.CreateAndCompare(wb =>
            {
                var ws = wb.AddWorksheet();
                ws.Range("A1:B2").FormulaArrayA1 = "1+2";
            }, @"Other\Formulas\ArrayFormula.xlsx");
        }

        [Test]
        public void ArrayFormulaCanBeLoaded()
        {
            TestHelper.LoadAndAssert(wb =>
            {
                var ws = wb.Worksheets.First();

                foreach (var arrayFormulaCell in ws.Range("A1:B2").Cells())
                {
                    Assert.AreEqual("1+2", arrayFormulaCell.FormulaA1);
                    Assert.AreEqual("A1:B2", arrayFormulaCell.FormulaReference.ToStringRelative());
                }

                var outsideCell = ws.Cell("A3");
                Assert.IsEmpty(outsideCell.FormulaA1);
                Assert.Null(outsideCell.FormulaReference);
            }, @"Other\Formulas\ArrayFormula.xlsx");
        }

        [Test]
        public void CanBeOnlyForOneCell()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var oneCell = ws.Cell("B3");

            oneCell.AsRange().FormulaArrayA1 = "2+5";

            Assert.True(oneCell.HasArrayFormula);
            Assert.AreEqual("2+5", oneCell.FormulaA1);
            Assert.AreEqual("B3:B3", oneCell.FormulaReference.ToStringRelative());
        }

        [TestCase("B2:C3")]
        [TestCase("B2:C4")]
        [TestCase("A1:D7")]
        public void SettingValueToContainingRangeClearsArrayFormula(string containingRange)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var arrayFormulaRange = ws.Range("B2:C3");
            arrayFormulaRange.FormulaArrayA1 = "5";

            ws.Range(containingRange).Value = Blank.Value;

            foreach (var cell in arrayFormulaRange.Cells())
            {
                Assert.AreEqual(Blank.Value, cell.Value);
                Assert.False(cell.HasArrayFormula);
                Assert.IsEmpty(cell.FormulaA1);
                Assert.Null(cell.FormulaReference);
            }
        }

        [TestCase("B2:D3")]
        [TestCase("A1:E4")]
        public void SettingFormulaToContainingRangeClearsOriginalArrayFormula(string overlapRange)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B2:D3").FormulaArrayA1 = "1";

            Assert.DoesNotThrow(() => ws.Range(overlapRange).FormulaArrayA1 = "2");
        }

        [TestCase("B2:B2")]
        [TestCase("B2:B3")]
        [TestCase("A1:C3")]
        [TestCase("D2:F3")]
        [TestCase("C:C")]
        [TestCase("2:2")]
        public void ArrayFormulaCantPartiallyOverlapWithAnotherArrayFormula(string partialOverlapRange)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B2:D3").FormulaArrayA1 = "1";

            Assert.That(() => ws.Range(partialOverlapRange).FormulaArrayA1 = "2",
                Throws.TypeOf<InvalidOperationException>()
                    .With.Message.EqualTo("Can't create array function that partially covers another array function."));
        }

        [TestCase("A1:B2")]
        [TestCase("A2")]
        public void ArrayFormulaCantOverlapWithMergedRange(string partialOverlapRange)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("A1:A2").Merge();

            Assert.That(() => ws.Range(partialOverlapRange).FormulaArrayA1 = "1",
                Throws.TypeOf<InvalidOperationException>()
                    .With.Message.EqualTo("Can't create array function over a merged range."));
        }

        [TestCase("A1:B2")]
        [TestCase("A1:C1")]
        public void ArrayFormulaCantOverlapWithTable(string formulaRange)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = "Name";
            ws.Cell("A2").Value = 5;
            ws.Range("A1:A2").CreateTable();

            Assert.That(() => ws.Range(formulaRange).FormulaArrayA1 = "1",
                Throws.TypeOf<InvalidOperationException>()
                    .With.Message.EqualTo("Can't create array function over a table."));
        }

        [Test]
        public void SettingArrayFormulaInvalidatesCells()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            Assert.False(ws.Cell("A1").NeedsRecalculation);
            Assert.False(ws.Cell("A2").NeedsRecalculation);

            ws.Range("A1:A2").FormulaArrayA1 = "ABS(-3)";

            Assert.True(ws.Cell("A1").NeedsRecalculation);
            Assert.True(ws.Cell("A2").NeedsRecalculation);
        }

        [Test]
        public void ReferencingItselfIsCircularError()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").FormulaA1 = "A2";
            ws.Range("A2").FormulaArrayA1 = "A1";

            Assert.That(() => _ = ws.Cell("A2").Value,
                Throws.TypeOf<InvalidOperationException>()
                    .With.Message.EqualTo("Formula in a cell '$Sheet1'!$A1 is part of a cycle."));
        }
    }
}
