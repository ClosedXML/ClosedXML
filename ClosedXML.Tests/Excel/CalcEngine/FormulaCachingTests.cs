﻿using ClosedXML.Excel;
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
            using (var wb = new XLWorkbook())
            {
                var sheet = wb.Worksheets.Add("TestSheet");
                var cell = sheet.Cell(1, 1);
                cell.Value = "1234567";

                Assert.IsFalse(cell.NeedsRecalculation);
            }
        }

        [Test]
        public void EditCellInvalidatesDependentCells()
        {
            using (var wb = new XLWorkbook())
            {
                var sheet = wb.Worksheets.Add("TestSheet");
                var cell = sheet.Cell(1, 1);
                var dependentCell = sheet.Cell(2, 1);
                dependentCell.FormulaA1 = "=A1";
                var _ = dependentCell.Value;

                cell.Value = "1234567";

                Assert.IsTrue(dependentCell.NeedsRecalculation);
            }
        }

        [Test]
        public void EditFormulaA1InvalidatesDependentCells()
        {
            using (var wb = new XLWorkbook())
            {
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

                Assert.AreEqual(15 + 150 + 1500, res1);
                Assert.AreEqual(15 + 300 + 3000, res2);
            }
        }

        [Test]
        public void EditFormulaR1C1InvalidatesDependentCells()
        {
            using (var wb = new XLWorkbook())
            {
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

                Assert.AreEqual(15 + 150 + 1500, res1);
                Assert.AreEqual(15 + 30 + 300, res2);
            }
        }

        [Test]
        public void InsertRowInvalidatesValues()
        {
            using (var wb = new XLWorkbook())
            {
                var sheet = wb.Worksheets.Add("TestSheet");
                var a4 = sheet.Cell("A4");
                a4.FormulaA1 = "=COUNTBLANK(A1:A3)";

                Assert.AreEqual(3, a4.Value);

                sheet.Row(2).InsertRowsAbove(2);

                Assert.AreEqual(5, sheet.Cell("A6").Value);
            }
        }

        [Test]
        public void DeleteRowModifiesFormulaAndInvalidatesValues()
        {
            using (var wb = new XLWorkbook())
            {
                var sheet = wb.Worksheets.Add("TestSheet");
                var original = sheet.Cell("A4");
                original.FormulaA1 = "=COUNTBLANK(A1:A3)";

                Assert.AreEqual(3, original.Value);

                sheet.Row(2).Delete();

                var shifted = sheet.Cell("A3");
                Assert.AreEqual("COUNTBLANK(A1:A2)", shifted.FormulaA1);
                Assert.AreEqual(2, shifted.Value);
            }
        }

        [Test]
        public void ChainedCalculationPreservesIntermediateValues()
        {
            using (var wb = new XLWorkbook())
            {
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

                Assert.AreEqual(15 + 150 + 1500, res);
                Assert.IsFalse(a4.NeedsRecalculation);
                Assert.IsFalse(a3.NeedsRecalculation);
                Assert.IsFalse(a2.NeedsRecalculation);
                Assert.AreEqual(150, a2.CachedValue);
                Assert.AreEqual(1500, a3.CachedValue);
                Assert.AreEqual(15 + 150 + 1500, a4.CachedValue);
            }
        }

        [Test]
        public void EditingAffectsDependentCells()
        {
            using (var wb = new XLWorkbook())
            {
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

                Assert.AreEqual(15 + 150 + 1500, res1);
                Assert.AreEqual(20 + 200 + 2000, res2);
            }
        }

        [Test]
        [TestCase("C4", new string[] { "C5" })]
        [TestCase("D4", new string[] { })]
        [TestCase("A1", new string[] { "A2", "A3", "A4", "C1", "C2", "C3", "C5" })]
        [TestCase("B2", new string[] { "B3", "B4", "C2", "C3", "C5" })]
        [TestCase("C2", new string[] { "C5" })]
        public void EditingDoesNotAffectNonDependingCells(string changedCell, string[] affectedCells)
        {
            using (var wb = new XLWorkbook())
            {
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

                Assert.AreEqual(affectedCells.Length, modifiedCells.Count());
                foreach (var cellAddress in affectedCells)
                {
                    Assert.IsTrue(modifiedCells.Any(cell => cell.Address.ToString() == cellAddress),
                        string.Format("Cell {0} is expected to need recalculation, but it does not", cellAddress));
                }
            }
        }

        [Test]
        public void CircularReferenceFailsCalculating()
        {
            using (var wb = new XLWorkbook())
            {
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
        }

        [Test]
        public void CircularReferenceRecalculationNeededDoesNotFail()
        {
            using (var wb = new XLWorkbook())
            {
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

                Assert.IsTrue(recalcNeededA1);
                Assert.IsTrue(recalcNeededA2);
                Assert.IsTrue(recalcNeededA3);
                Assert.IsTrue(recalcNeededA4);
            }
        }

        [Test]
        public void DeleteWorksheetInvalidatesValues()
        {
            using (var wb = new XLWorkbook())
            {
                var sheet1 = wb.Worksheets.Add("Sheet1");
                var sheet2 = wb.Worksheets.Add("Sheet2");
                var sheet1_a1 = sheet1.Cell("A1");
                var sheet2_a1 = sheet2.Cell("A1");
                sheet1_a1.FormulaA1 = "Sheet2!A1";
                sheet2_a1.Value = "TestValue";

                var valueBeforeDeletion = sheet1_a1.Value;
                sheet2.Delete();
                var valueAfterDeletion = sheet1_a1.Value;

                Assert.AreEqual("TestValue", valueBeforeDeletion);
                Assert.AreEqual(XLError.CellReference, valueAfterDeletion);
            }
        }

        [Test]
        public void CachedValueToExternalWorkbook()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\ExternalLinks\WorkbookWithExternalLink.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheets.First();
                var cell = ws.Cell("B2");
                Assert.IsFalse(cell.NeedsRecalculation);
                Assert.IsTrue(cell.HasFormula);

                // This will fail when we start supporting external links
                Assert.IsTrue(cell.FormulaA1.StartsWith("[1]"));

                Assert.AreEqual("hello world", cell.CachedValue);
                Assert.AreEqual("hello world", cell.Value);

                Assert.AreEqual(11, ws.Evaluate("LEN(B2)"));

                Assert.Throws(Is.TypeOf<NotImplementedException>().And.Message.EqualTo("References from other files are not yet implemented."), () => wb.RecalculateAllFormulas());
            }
        }

        [Test]
        public void ChangingValueChangesCachedValue()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Test");
                var cell = ws.Cell(1, 1);

                cell.Value = "Hello";
                Assert.AreEqual("Hello", cell.CachedValue);

                cell.Value = 74.0;
                Assert.AreEqual(74.0, cell.CachedValue);

                cell.Value = new DateTime(2019, 1, 1, 14, 0, 0);
                Assert.AreEqual(new DateTime(2019, 1, 1, 14, 0, 0), cell.CachedValue);
            }
        }
    }
}
