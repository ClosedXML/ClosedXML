using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Linq;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    public class FormulaCachingTests
    {
        [Test]
        public void StaticCellDoesNotNeedRecalculation()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.Worksheets.Add("TestSheet");
            var cell = sheet.Cell(1, 1);
            cell.Value = "1234567";

            Assert.That(cell.NeedsRecalculation, Is.False);
        }

        [Test]
        public void EditCellInvalidatesDependentCells()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.Worksheets.Add("TestSheet");
            var cell = sheet.Cell(1, 1);
            var dependentCell = sheet.Cell(2, 1);
            dependentCell.FormulaA1 = "=A1";
            var _ = dependentCell.Value;

            cell.Value = "1234567";

            Assert.That(dependentCell.NeedsRecalculation, Is.True);
        }

        [Test]
        public void EditFormulaA1InvalidatesDependentCells()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.Worksheets.Add("TestSheet");
            var a1 = sheet.Cell("A1");
            var a2 = sheet.Cell("A2");
            var a3 = sheet.Cell("A3");
            var a4 = sheet.Cell("A4");
            a2.FormulaA1 = "=A1*10";
            a3.FormulaA1 = "=A2*10";
            a4.FormulaA1 = "=SUM(A1:A3)";
            a1.Value = 15;

            var res1 = a4.Value;
            a2.FormulaA1 = "=A1*20";
            var res2 = a4.Value;

            Assert.Multiple(() =>
            {
                Assert.That(res1, Is.EqualTo(15 + 150 + 1500));
                Assert.That(res2, Is.EqualTo(15 + 300 + 3000));
            });
        }

        [Test]
        public void EditFormulaR1C1InvalidatesDependentCells()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.Worksheets.Add("TestSheet");
            var a1 = sheet.Cell("A1");
            var a2 = sheet.Cell("A2");
            var a3 = sheet.Cell("A3");
            var a4 = sheet.Cell("A4");
            a2.FormulaA1 = "=A1*10";
            a3.FormulaA1 = "=A2*10";
            a4.FormulaA1 = "=SUM(A1:A3)";
            a1.Value = 15;

            var res1 = a4.Value;
            a2.FormulaR1C1 = "=R[-1]C*2";
            var res2 = a4.Value;

            Assert.Multiple(() =>
            {
                Assert.That(res1, Is.EqualTo(15 + 150 + 1500));
                Assert.That(res2, Is.EqualTo(15 + 30 + 300));
            });
        }

        [Test]
        public void InsertRowInvalidatesValues()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.Worksheets.Add("TestSheet");
            var a4 = sheet.Cell("A4");
            a4.FormulaA1 = "=COUNTBLANK(A1:A3)";

            Assert.That(a4.Value, Is.EqualTo(3));

            sheet.Row(2).InsertRowsAbove(2);

            Assert.That(sheet.Cell("A6").Value, Is.EqualTo(5));
        }

        [Test]
        public void DeleteRowModifiesFormulaAndInvalidatesValues()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.Worksheets.Add("TestSheet");
            var original = sheet.Cell("A4");
            original.FormulaA1 = "=COUNTBLANK(A1:A3)";

            Assert.That(original.Value, Is.EqualTo(3));

            sheet.Row(2).Delete();

            var shifted = sheet.Cell("A3");
            Assert.Multiple(() =>
            {
                Assert.That(shifted.FormulaA1, Is.EqualTo("COUNTBLANK(A1:A2)"));
                Assert.That(shifted.Value, Is.EqualTo(2));
            });
        }

        [Test]
        public void ChainedCalculationPreservesIntermediateValues()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.Worksheets.Add("TestSheet");
            var a1 = sheet.Cell("A1");
            var a2 = sheet.Cell("A2");
            var a3 = sheet.Cell("A3");
            var a4 = sheet.Cell("A4");
            a2.FormulaA1 = "=A1*10";
            a3.FormulaA1 = "=A2*10";
            a4.FormulaA1 = "=SUM(A1:A3)";

            a1.Value = 15;
            var res = a4.Value;

            Assert.Multiple(() =>
            {
                Assert.That(res, Is.EqualTo(15 + 150 + 1500));
                Assert.That(a4.NeedsRecalculation, Is.False);
                Assert.That(a3.NeedsRecalculation, Is.False);
                Assert.That(a2.NeedsRecalculation, Is.False);
                Assert.That(a2.CachedValue, Is.EqualTo(150));
                Assert.That(a3.CachedValue, Is.EqualTo(1500));
                Assert.That(a4.CachedValue, Is.EqualTo(15 + 150 + 1500));
            });
        }

        [Test]
        public void EditingAffectsDependentCells()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.Worksheets.Add("TestSheet");
            var a1 = sheet.Cell("A1");
            var a2 = sheet.Cell("A2");
            var a3 = sheet.Cell("A3");
            var a4 = sheet.Cell("A4");
            a2.FormulaA1 = "=A1*10";
            a3.FormulaA1 = "=A2*10";
            a4.FormulaA1 = "=SUM(A1:A3)";
            a1.Value = 15;

            var res1 = a4.Value;
            a1.Value = 20;
            var res2 = a4.Value;

            Assert.Multiple(() =>
            {
                Assert.That(res1, Is.EqualTo(15 + 150 + 1500));
                Assert.That(res2, Is.EqualTo(20 + 200 + 2000));
            });
        }

        [Test]
        [TestCase("C4", new string[] { "C5" })]
        [TestCase("D4", new string[] { })]
        [TestCase("A1", new string[] { "A2", "A3", "A4", "C1", "C2", "C3", "C5" })]
        [TestCase("B2", new string[] { "B3", "B4", "C2", "C3", "C5" })]
        [TestCase("C2", new string[] { "C5" })]
        public void EditingDoesNotAffectNonDependingCells(string changedCell, string[] affectedCells)
        {
            using var wb = new XLWorkbook();
            var sheet = wb.Worksheets.Add("TestSheet");
            sheet.Cell("A2").FormulaA1 = "A1+1";
            sheet.Cell("A3").FormulaA1 = "SUM(A1:A2)";
            sheet.Cell("A4").FormulaA1 = "SUM(A1:A3)";
            sheet.Cell("B2").FormulaA1 = "B1+1";
            sheet.Cell("B3").FormulaA1 = "SUM(B1:B2)";
            sheet.Cell("B4").FormulaA1 = "SUM(B1:B3)";
            sheet.Cell("C1").FormulaA1 = "SUM(A1:B1)";
            sheet.Cell("C2").FormulaA1 = "SUM(A2:B2)";
            sheet.Cell("C3").FormulaA1 = "SUM(A3:B3)";
            sheet.Cell("C5").FormulaA1 = "SUM($A$1:$C$4)";
            sheet.RecalculateAllFormulas();
            var allCells = sheet.CellsUsed();

            sheet.Cell(changedCell).Value = 100;
            var modifiedCells = allCells.Where(cell => cell.NeedsRecalculation);

            Assert.That(modifiedCells.Count(), Is.EqualTo(affectedCells.Length));
            foreach (var cellAddress in affectedCells)
            {
                Assert.That(modifiedCells.Any(cell => cell.Address.ToString() == cellAddress),
                    Is.True,
                    string.Format("Cell {0} is expected to need recalculation, but it does not", cellAddress));
            }
        }

        [Test]
        public void CircularReferenceFailsCalculating()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.Worksheets.Add("TestSheet");
            var a1 = sheet.Cell("A1");
            var a2 = sheet.Cell("A2");
            var a3 = sheet.Cell("A3");
            var a4 = sheet.Cell("A4");

            a2.FormulaA1 = "=A1*10";
            a3.FormulaA1 = "=A2*10";
            a4.FormulaA1 = "=A3*10";
            a1.FormulaA1 = "A2+A3+A4";

            var getValueA1 = new TestDelegate(() => { var v = a1.Value; });
            var getValueA2 = new TestDelegate(() => { var v = a2.Value; });
            var getValueA3 = new TestDelegate(() => { var v = a3.Value; });
            var getValueA4 = new TestDelegate(() => { var v = a4.Value; });

            Assert.Throws(typeof(InvalidOperationException), getValueA1);
            Assert.Throws(typeof(InvalidOperationException), getValueA2);
            Assert.Throws(typeof(InvalidOperationException), getValueA3);
            Assert.Throws(typeof(InvalidOperationException), getValueA4);
        }

        [Test]
        public void CircularReferenceRecalculationNeededDoesNotFail()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.Worksheets.Add("TestSheet");
            var a1 = sheet.Cell("A1");
            var a2 = sheet.Cell("A2");
            var a3 = sheet.Cell("A3");
            var a4 = sheet.Cell("A4");

            a2.FormulaA1 = "=A1*10";
            a3.FormulaA1 = "=A2*10";
            a4.FormulaA1 = "=A3*10";
            var _ = a4.Value;
            a1.FormulaA1 = "=SUM(A2:A4)";

            var recalcNeededA1 = a1.NeedsRecalculation;
            var recalcNeededA2 = a2.NeedsRecalculation;
            var recalcNeededA3 = a3.NeedsRecalculation;
            var recalcNeededA4 = a4.NeedsRecalculation;

            Assert.Multiple(() =>
            {
                Assert.That(recalcNeededA1, Is.True);
                Assert.That(recalcNeededA2, Is.True);
                Assert.That(recalcNeededA3, Is.True);
                Assert.That(recalcNeededA4, Is.True);
            });
        }

        [Test]
        public void DeleteWorksheetInvalidatesValues()
        {
            using var wb = new XLWorkbook();
            var sheet1 = wb.Worksheets.Add("Sheet1");
            var sheet2 = wb.Worksheets.Add("Sheet2");
            var sheet1_a1 = sheet1.Cell("A1");
            var sheet2_a1 = sheet2.Cell("A1");
            sheet1_a1.FormulaA1 = "Sheet2!A1";
            sheet2_a1.Value = "TestValue";

            var valueBeforeDeletion = sheet1_a1.Value;
            sheet2.Delete();
            var valueAfterDeletion = sheet1_a1.Value;

            Assert.Multiple(() =>
            {
                Assert.That(valueBeforeDeletion, Is.EqualTo("TestValue"));
                Assert.That(valueAfterDeletion, Is.EqualTo(XLError.CellReference));
            });
        }

        [Test]
        public void CachedValueToExternalWorkbook()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\ExternalLinks\WorkbookWithExternalLink.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheets.First();
            var cell = ws.Cell("B2");
            Assert.Multiple(() =>
            {
                Assert.That(cell.NeedsRecalculation, Is.False);
                Assert.That(cell.HasFormula, Is.True);

                // This will fail when we start supporting external links
                Assert.That(cell.FormulaA1, Does.StartWith("[1]"));

                Assert.That(cell.CachedValue, Is.EqualTo("hello world"));
                Assert.That(cell.Value, Is.EqualTo("hello world"));

                Assert.That(ws.Evaluate("LEN(B2)"), Is.EqualTo(11));
            });

            Assert.Throws(Is.TypeOf<NotImplementedException>().And.Message.EqualTo("References from other files are not yet implemented."), () => wb.RecalculateAllFormulas());
        }

        [Test]
        public void ChangingValueChangesCachedValue()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Test");
            var cell = ws.Cell(1, 1);

            cell.Value = "Hello";
            Assert.That(cell.CachedValue, Is.EqualTo("Hello"));

            cell.Value = 74.0;
            Assert.That(cell.CachedValue, Is.EqualTo(74.0));

            cell.Value = new DateTime(2019, 1, 1, 14, 0, 0);
            Assert.That(cell.CachedValue, Is.EqualTo(new DateTime(2019, 1, 1, 14, 0, 0)));
        }
    }
}
