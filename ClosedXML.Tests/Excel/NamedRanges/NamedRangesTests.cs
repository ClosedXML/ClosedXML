// Keep this file CodeMaid organised and cleaned
using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.IO;
using System.Linq;

namespace ClosedXML.Tests.Excel
{
    [TestFixture]
    public class NamedRangesTests
    {
        [Test]
        public void CanEvaluateNamedMultiRange()
        {
            using (var wb = new XLWorkbook())
            {
                var ws1 = wb.AddWorksheet("Sheet1");
                ws1.Range("A1:C1").Value = 1;
                ws1.Range("A3:C3").Value = 3;
                wb.NamedRanges.Add("TEST", ws1.Ranges("A1:C1,A3:C3"));

                ws1.Cell(2, 1).FormulaA1 = "=SUM(TEST)";

                Assert.AreEqual(12.0, (double)ws1.Cell(2, 1).Value, XLHelper.Epsilon);
            }
        }

        [Test]
        public void CanGetNamedFromAnother()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("Sheet1");
            ws1.Cell("A1").SetValue(1).AddToNamed("value1");

            Assert.AreEqual(1, wb.Cell("value1").GetValue<int>());
            Assert.AreEqual(1, wb.Range("value1").FirstCell().GetValue<int>());

            Assert.AreEqual(1, ws1.Cell("value1").GetValue<int>());
            Assert.AreEqual(1, ws1.Range("value1").FirstCell().GetValue<int>());

            var ws2 = wb.Worksheets.Add("Sheet2");

            ws2.Cell("A1").SetFormulaA1("=value1").AddToNamed("value2");

            Assert.AreEqual(1, wb.Cell("value2").GetValue<int>());
            Assert.AreEqual(1, wb.Range("value2").FirstCell().GetValue<int>());

            Assert.AreEqual(1, ws2.Cell("value1").GetValue<int>());
            Assert.AreEqual(1, ws2.Range("value1").FirstCell().GetValue<int>());

            Assert.AreEqual(1, ws2.Cell("value2").GetValue<int>());
            Assert.AreEqual(1, ws2.Range("value2").FirstCell().GetValue<int>());
        }

        [Test]
        public void CanGetValidNamedRanges()
        {
            using (var wb = new XLWorkbook())
            {
                var ws1 = wb.Worksheets.Add("Sheet 1");
                var ws2 = wb.Worksheets.Add("Sheet 2");
                var ws3 = wb.Worksheets.Add("Sheet'3");

                ws1.Range("A1:D1").AddToNamed("Named range 1", XLScope.Worksheet);
                ws1.Range("A2:D2").AddToNamed("Named range 2", XLScope.Workbook);
                ws2.Range("A3:D3").AddToNamed("Named range 3", XLScope.Worksheet);
                ws2.Range("A4:D4").AddToNamed("Named range 4", XLScope.Workbook);
                wb.NamedRanges.Add("Named range 5", new XLRanges
                {
                    ws1.Range("A5:D5"),
                    ws3.Range("A5:D5")
                });

                ws2.Delete();
                ws3.Delete();

                var globalValidRanges = wb.NamedRanges.ValidNamedRanges();
                var globalInvalidRanges = wb.NamedRanges.InvalidNamedRanges();
                var localValidRanges = ws1.NamedRanges.ValidNamedRanges();
                var localInvalidRanges = ws1.NamedRanges.InvalidNamedRanges();

                Assert.AreEqual(1, globalValidRanges.Count());
                Assert.AreEqual("Named range 2", globalValidRanges.First().Name);

                Assert.AreEqual(2, globalInvalidRanges.Count());
                Assert.AreEqual("Named range 4", globalInvalidRanges.First().Name);
                Assert.AreEqual("Named range 5", globalInvalidRanges.Last().Name);

                Assert.AreEqual(1, localValidRanges.Count());
                Assert.AreEqual("Named range 1", localValidRanges.First().Name);

                Assert.AreEqual(0, localInvalidRanges.Count());
            }
        }

        [Test]
        public void CanRenameNamedRange()
        {
            using (var wb = new XLWorkbook())
            {
                var ws1 = wb.AddWorksheet("Sheet1");
                var nr1 = wb.NamedRanges.Add("TEST", "=0.1");

                Assert.IsTrue(wb.NamedRanges.TryGetValue("TEST", out IXLNamedRange _));
                Assert.IsFalse(wb.NamedRanges.TryGetValue("TEST1", out IXLNamedRange _));

                nr1.Name = "TEST1";

                Assert.IsFalse(wb.NamedRanges.TryGetValue("TEST", out IXLNamedRange _));
                Assert.IsTrue(wb.NamedRanges.TryGetValue("TEST1", out IXLNamedRange _));

                var nr2 = wb.NamedRanges.Add("TEST2", "=TEST1*2");

                ws1.Cell(1, 1).FormulaA1 = "TEST1";
                ws1.Cell(2, 1).FormulaA1 = "TEST1*10";
                ws1.Cell(3, 1).FormulaA1 = "TEST2";
                ws1.Cell(4, 1).FormulaA1 = "TEST2*3";

                Assert.AreEqual(0.1, (double)ws1.Cell(1, 1).Value, XLHelper.Epsilon);
                Assert.AreEqual(1.0, (double)ws1.Cell(2, 1).Value, XLHelper.Epsilon);
                Assert.AreEqual(0.2, (double)ws1.Cell(3, 1).Value, XLHelper.Epsilon);
                Assert.AreEqual(0.6, (double)ws1.Cell(4, 1).Value, XLHelper.Epsilon);
            }
        }

        [Test]
        public void CanSaveAndLoadNamedRanges()
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = new XLWorkbook())
                {
                    var sheet1 = wb.Worksheets.Add("Sheet1");
                    var sheet2 = wb.Worksheets.Add("Sheet2");

                    wb.NamedRanges.Add("wbNamedRange",
                        "Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G");
                    sheet1.NamedRanges.Add("sheet1NamedRange",
                        "Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G");
                    sheet2.NamedRanges.Add("sheet2NamedRange", "Sheet1!A1,Sheet2!A1");

                    wb.SaveAs(ms);
                }

                using (var wb = new XLWorkbook(ms))
                {
                    var sheet1 = wb.Worksheet("Sheet1");
                    var sheet2 = wb.Worksheet("Sheet2");

                    Assert.AreEqual(1, wb.NamedRanges.Count());
                    Assert.AreEqual("wbNamedRange", wb.NamedRanges.Single().Name);
                    Assert.AreEqual("Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G", wb.NamedRanges.Single().RefersTo);
                    Assert.AreEqual(5, wb.NamedRanges.Single().Ranges.Count);

                    Assert.AreEqual(1, sheet1.NamedRanges.Count());
                    Assert.AreEqual("sheet1NamedRange", sheet1.NamedRanges.Single().Name);
                    Assert.AreEqual("Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G", sheet1.NamedRanges.Single().RefersTo);
                    Assert.AreEqual(5, sheet1.NamedRanges.Single().Ranges.Count);

                    Assert.AreEqual(1, sheet2.NamedRanges.Count());
                    Assert.AreEqual("sheet2NamedRange", sheet2.NamedRanges.Single().Name);
                    Assert.AreEqual("Sheet1!A1,Sheet2!A1", sheet2.NamedRanges.Single().RefersTo);
                    Assert.AreEqual(2, sheet2.NamedRanges.Single().Ranges.Count);
                }
            }
        }

        [Test]
        public void CopyNamedRangeDifferentWorksheets()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("Sheet1");
            var ws2 = wb.Worksheets.Add("Sheet2");
            var ranges = new XLRanges();
            ranges.Add(ws1.Range("B2:E6"));
            ranges.Add(ws2.Range("D1:E2"));
            var original = ws1.NamedRanges.Add("Named range", ranges);

            var copy = original.CopyTo(ws2);

            Assert.AreEqual(1, ws1.NamedRanges.Count());
            Assert.AreEqual(1, ws2.NamedRanges.Count());
            Assert.AreEqual(2, original.Ranges.Count);
            Assert.AreEqual(2, copy.Ranges.Count);
            Assert.AreEqual(original.Name, copy.Name);
            Assert.AreEqual(original.Scope, copy.Scope);
            Assert.AreEqual("Sheet1!B2:E6", original.Ranges.First().RangeAddress.ToString(XLReferenceStyle.A1, true));
            Assert.AreEqual("Sheet2!D1:E2", original.Ranges.Last().RangeAddress.ToString(XLReferenceStyle.A1, true));
            Assert.AreEqual("Sheet2!D1:E2", copy.Ranges.First().RangeAddress.ToString(XLReferenceStyle.A1, true));
            Assert.AreEqual("Sheet2!B2:E6", copy.Ranges.Last().RangeAddress.ToString(XLReferenceStyle.A1, true));
        }

        [Test]
        public void CopyNamedRangeSameWorksheet()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("Sheet1");
            ws1.Range("B2:E6").AddToNamed("Named range", XLScope.Worksheet);
            var nr = ws1.NamedRange("Named range");

            TestDelegate action = () => nr.CopyTo(ws1);

            Assert.Throws(typeof(InvalidOperationException), action);
        }

        [Test]
        public void DeleteColumnUsedInNamedRange()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.FirstCell().SetValue("Column1");
                ws.FirstCell().CellRight().SetValue("Column2").Style.Font.SetBold();
                ws.FirstCell().CellRight(2).SetValue("Column3");
                ws.NamedRanges.Add("MyRange", "A1:C1");

                ws.Column(1).Delete();

                Assert.IsTrue(ws.Cell("A1").Style.Font.Bold);
                Assert.AreEqual("Column3", ws.Cell("B1").GetValue<string>());
                Assert.IsEmpty(ws.Cell("C1").GetValue<string>());
            }
        }

        [Test]
        public void MovingRanges()
        {
            var wb = new XLWorkbook();

            IXLWorksheet sheet1 = wb.Worksheets.Add("Sheet1");
            IXLWorksheet sheet2 = wb.Worksheets.Add("Sheet2");

            wb.NamedRanges.Add("wbNamedRange",
                "Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G");
            sheet1.NamedRanges.Add("sheet1NamedRange",
                "Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G");
            sheet2.NamedRanges.Add("sheet2NamedRange", "Sheet1!A1,Sheet2!A1");

            sheet1.Row(1).InsertRowsAbove(2);
            sheet1.Row(1).Delete();
            sheet1.Column(1).InsertColumnsBefore(2);
            sheet1.Column(1).Delete();

            Assert.AreEqual("Sheet1!$C$3,Sheet1!$C$4:$D$4,Sheet2!$D$3:$D$4,Sheet1!$7:$8,Sheet1!$G:$H",
                wb.NamedRanges.First().RefersTo);
            Assert.AreEqual("Sheet1!$C$3,Sheet1!$C$4:$D$4,Sheet2!$D$3:$D$4,Sheet1!$7:$8,Sheet1!$G:$H",
                sheet1.NamedRanges.First().RefersTo);
            Assert.AreEqual("Sheet1!B2,Sheet2!A1", sheet2.NamedRanges.First().RefersTo);

            wb.NamedRanges.ForEach(nr => Assert.AreEqual(XLNamedRangeScope.Workbook, nr.Scope));
            sheet1.NamedRanges.ForEach(nr => Assert.AreEqual(XLNamedRangeScope.Worksheet, nr.Scope));
            sheet2.NamedRanges.ForEach(nr => Assert.AreEqual(XLNamedRangeScope.Worksheet, nr.Scope));
        }

        [Test, Ignore("Muted until shifting is fixed (see #880)")]
        public void NamedRangeBecomesInvalidOnRangeAndWorksheetDeleting()
        {
            using (var wb = new XLWorkbook())
            {
                var ws1 = wb.Worksheets.Add("Sheet 1");
                var ws2 = wb.Worksheets.Add("Sheet 2");
                ws1.Range("A1:B2").AddToNamed("Simple", XLScope.Workbook);
                wb.NamedRanges.Add("Compound", new XLRanges
                {
                    ws1.Range("C1:D2"),
                    ws2.Range("A10:D15")
                });

                ws1.Rows(1, 5).Delete();
                ws1.Delete();

                Assert.AreEqual(2, wb.NamedRanges.Count());
                Assert.AreEqual(0, wb.NamedRanges.ValidNamedRanges().Count());
                Assert.AreEqual("#REF!#REF!", wb.NamedRanges.ElementAt(0).RefersTo);
                Assert.AreEqual("#REF!#REF!,'Sheet 2'!A10:D15", wb.NamedRanges.ElementAt(0).RefersTo);
            }
        }

        [Test, Ignore("Muted until shifting is fixed (see #880)")]
        public void NamedRangeBecomesInvalidOnRangeDeleting()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet 1");
                ws.Range("A1:B2").AddToNamed("Simple", XLScope.Workbook);
                wb.NamedRanges.Add("Compound", new XLRanges
                {
                    ws.Range("C1:D2"),
                    ws.Range("A10:D15")
                });

                ws.Rows(1, 5).Delete();

                Assert.AreEqual(2, wb.NamedRanges.Count());
                Assert.AreEqual(0, wb.NamedRanges.ValidNamedRanges().Count());
                Assert.AreEqual("'Sheet 1'!#REF!", wb.NamedRanges.ElementAt(0).RefersTo);
                Assert.AreEqual("'Sheet 1'!#REF!,'Sheet 1'!A5:D10", wb.NamedRanges.ElementAt(0).RefersTo);
            }
        }

        [Test]
        public void NamedRangeMayReferToExpression()
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = new XLWorkbook())
                {
                    var ws1 = wb.AddWorksheet("Sheet1");
                    wb.NamedRanges.Add("TEST", "=0.1");
                    wb.NamedRanges.Add("TEST2", "=TEST*2");

                    ws1.Cell(1, 1).FormulaA1 = "TEST";
                    ws1.Cell(2, 1).FormulaA1 = "TEST*10";
                    ws1.Cell(3, 1).FormulaA1 = "TEST2";
                    ws1.Cell(4, 1).FormulaA1 = "TEST2*3";

                    Assert.AreEqual(0.1, (double)ws1.Cell(1, 1).Value, XLHelper.Epsilon);
                    Assert.AreEqual(1.0, (double)ws1.Cell(2, 1).Value, XLHelper.Epsilon);
                    Assert.AreEqual(0.2, (double)ws1.Cell(3, 1).Value, XLHelper.Epsilon);
                    Assert.AreEqual(0.6, (double)ws1.Cell(4, 1).Value, XLHelper.Epsilon);

                    wb.SaveAs(ms);
                }

                using (var wb = new XLWorkbook(ms))
                {
                    var ws1 = wb.Worksheets.First();

                    Assert.AreEqual(0.1, (double)ws1.Cell(1, 1).Value, XLHelper.Epsilon);
                    Assert.AreEqual(1.0, (double)ws1.Cell(2, 1).Value, XLHelper.Epsilon);
                    Assert.AreEqual(0.2, (double)ws1.Cell(3, 1).Value, XLHelper.Epsilon);
                    Assert.AreEqual(0.6, (double)ws1.Cell(4, 1).Value, XLHelper.Epsilon);
                }
            }
        }

        [Test]
        public void NamedRangeReferringToMultipleRangesCanBeSavedAndLoaded()
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = new XLWorkbook())
                {
                    var ws = wb.Worksheets.Add("Sheet 1");

                    wb.NamedRanges.Add("Multirange named range", new XLRanges
                    {
                        ws.Range("A5:D5"),
                        ws.Range("A15:D15")
                    });

                    wb.SaveAs(ms);
                }

                using (var wb = new XLWorkbook(ms))
                {
                    Assert.AreEqual(1, wb.NamedRanges.Count());
                    var nr = wb.NamedRanges.Single() as XLNamedRange;
                    Assert.AreEqual("'Sheet 1'!$A$5:$D$5,'Sheet 1'!$A$15:$D$15", nr.RefersTo);
                    Assert.AreEqual(2, nr.Ranges.Count);
                    Assert.AreEqual("'Sheet 1'!A5:D5", nr.Ranges.First().RangeAddress.ToString(XLReferenceStyle.A1, true));
                    Assert.AreEqual("'Sheet 1'!A15:D15", nr.Ranges.Last().RangeAddress.ToString(XLReferenceStyle.A1, true));
                    Assert.AreEqual(2, nr.RangeList.Count);
                    Assert.AreEqual("'Sheet 1'!$A$5:$D$5", nr.RangeList.First());
                    Assert.AreEqual("'Sheet 1'!$A$15:$D$15", nr.RangeList.Last());
                }
            }
        }

        [Test]
        public void NamedRangesBecomeInvalidOnWorksheetDeleting()
        {
            using (var wb = new XLWorkbook())
            {
                var ws1 = wb.Worksheets.Add("Sheet 1");
                var ws2 = wb.Worksheets.Add("Sheet 2");
                var ws3 = wb.Worksheets.Add("Sheet'3");

                ws1.Range("A1:D1").AddToNamed("Named range 1", XLScope.Worksheet);
                ws1.Range("A2:D2").AddToNamed("Named range 2", XLScope.Workbook);
                ws2.Range("A3:D3").AddToNamed("Named range 3", XLScope.Worksheet);
                ws2.Range("A4:D4").AddToNamed("Named range 4", XLScope.Workbook);
                wb.NamedRanges.Add("Named range 5", new XLRanges
                {
                    ws1.Range("A5:D5"),
                    ws3.Range("A5:D5")
                });

                ws2.Delete();
                ws3.Delete();

                Assert.AreEqual(1, ws1.NamedRanges.Count());
                Assert.AreEqual("Named range 1", ws1.NamedRanges.First().Name);
                Assert.AreEqual(XLNamedRangeScope.Worksheet, ws1.NamedRanges.First().Scope);
                Assert.AreEqual("'Sheet 1'!$A$1:$D$1", ws1.NamedRanges.First().RefersTo);
                Assert.AreEqual("'Sheet 1'!A1:D1", ws1.NamedRanges.First().Ranges.Single().RangeAddress.ToString(XLReferenceStyle.A1, true));

                Assert.AreEqual(3, wb.NamedRanges.Count());

                Assert.AreEqual("Named range 2", wb.NamedRanges.ElementAt(0).Name);
                Assert.AreEqual(XLNamedRangeScope.Workbook, wb.NamedRanges.ElementAt(0).Scope);
                Assert.AreEqual("'Sheet 1'!$A$2:$D$2", wb.NamedRanges.ElementAt(0).RefersTo);
                Assert.AreEqual("'Sheet 1'!A2:D2", wb.NamedRanges.ElementAt(0).Ranges.Single().RangeAddress.ToString(XLReferenceStyle.A1, true));

                Assert.AreEqual("Named range 4", wb.NamedRanges.ElementAt(1).Name);
                Assert.AreEqual(XLNamedRangeScope.Workbook, wb.NamedRanges.ElementAt(1).Scope);
                Assert.AreEqual("#REF!$A$4:$D$4", wb.NamedRanges.ElementAt(1).RefersTo);
                Assert.IsFalse(wb.NamedRanges.ElementAt(1).Ranges.Any());

                Assert.AreEqual("Named range 5", wb.NamedRanges.ElementAt(2).Name);
                Assert.AreEqual(XLNamedRangeScope.Workbook, wb.NamedRanges.ElementAt(2).Scope);
                Assert.AreEqual("'Sheet 1'!$A$5:$D$5,#REF!$A$5:$D$5", wb.NamedRanges.ElementAt(2).RefersTo);
                Assert.AreEqual(1, wb.NamedRanges.ElementAt(2).Ranges.Count);
                Assert.AreEqual("'Sheet 1'!A5:D5", wb.NamedRanges.ElementAt(2).Ranges.Single().RangeAddress.ToString(XLReferenceStyle.A1, true));
            }
        }

        [Test]
        public void NamedRangesFromDeletedSheetAreSavedWithoutAddress()
        {
            // Range address referring to the deleted sheet look like #REF!A1:B2.
            // But workbooks with such references in named ranges Excel considers as broken files.
            // It requires #REF!

            using (var ms = new MemoryStream())
            {
                using (var wb = new XLWorkbook())
                {
                    wb.Worksheets.Add("Sheet 1");
                    var ws2 = wb.Worksheets.Add("Sheet 2");
                    ws2.Range("A4:D4").AddToNamed("Test named range", XLScope.Workbook);
                    ws2.Delete();
                    wb.SaveAs(ms);
                }

                using (var wb = new XLWorkbook(ms))
                {
                    Assert.AreEqual("#REF!", wb.NamedRanges.Single().RefersTo);
                }
            }
        }

        [Test]
        public void NamedRangesWhenCopyingWorksheets()
        {
            using (var wb = new XLWorkbook())
            {
                var ws1 = wb.AddWorksheet("Sheet1");
                ws1.FirstCell().Value = Enumerable.Range(1, 10);
                wb.NamedRanges.Add("wbNamedRange", ws1.Range("A1:A10"));
                ws1.NamedRanges.Add("wsNamedRange", ws1.Range("A3"));

                var ws2 = wb.AddWorksheet("Sheet2");
                ws2.FirstCell().Value = Enumerable.Range(101, 10);
                ws1.NamedRanges.Add("wsNamedRangeAcrossSheets", ws2.Range("A4"));

                ws1.Cell("C1").FormulaA1 = "=wbNamedRange";
                ws1.Cell("C2").FormulaA1 = "=wsNamedRange";
                ws1.Cell("C3").FormulaA1 = "=wsNamedRangeAcrossSheets";

                Assert.AreEqual(1, ws1.Cell("C1").Value);
                Assert.AreEqual(3, ws1.Cell("C2").Value);
                Assert.AreEqual(104, ws1.Cell("C3").Value);

                var wsCopy = ws1.CopyTo("Copy");
                Assert.AreEqual(1, wsCopy.Cell("C1").Value);
                Assert.AreEqual(3, wsCopy.Cell("C2").Value);
                Assert.AreEqual(104, wsCopy.Cell("C3").Value);

                Assert.AreEqual("Sheet1!A1:A10",
                    wb.NamedRange("wbNamedRange").Ranges.First().RangeAddress.ToStringRelative(true));
                Assert.AreEqual("Copy!A3:A3",
                    wsCopy.NamedRange("wsNamedRange").Ranges.First().RangeAddress.ToStringRelative(true));
                Assert.AreEqual("Sheet2!A4:A4",
                    wsCopy.NamedRange("wsNamedRangeAcrossSheets").Ranges.First().RangeAddress.ToStringRelative(true));
            }
        }

        [Test]
        public void SavedNamedRangesBecomeInvalidOnWorksheetDeleting()
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = new XLWorkbook())
                {
                    var ws1 = wb.Worksheets.Add("Sheet 1");
                    var ws2 = wb.Worksheets.Add("Sheet2");
                    var ws3 = wb.Worksheets.Add("Sheet'3");

                    ws1.Range("A1:D1").AddToNamed("Named range 1", XLScope.Worksheet);
                    ws1.Range("A2:D2").AddToNamed("Named range 2", XLScope.Workbook);
                    ws2.Range("A3:D3").AddToNamed("Named range 3", XLScope.Worksheet);
                    ws2.Range("A4:D4").AddToNamed("Named range 4", XLScope.Workbook);
                    wb.NamedRanges.Add("Named range 5", new XLRanges
                    {
                        ws1.Range("A5:D5"),
                        ws3.Range("A5:D5")
                    });

                    wb.SaveAs(ms);
                }

                using (var wb = new XLWorkbook(ms))
                {
                    wb.Worksheet("Sheet2").Delete();
                    wb.Worksheet("Sheet'3").Delete();
                    wb.Save();
                }

                using (var wb = new XLWorkbook(ms))
                {
                    var ws1 = wb.Worksheet("Sheet 1");
                    Assert.AreEqual(1, ws1.NamedRanges.Count());
                    Assert.AreEqual("Named range 1", ws1.NamedRanges.First().Name);
                    Assert.AreEqual(XLNamedRangeScope.Worksheet, ws1.NamedRanges.First().Scope);
                    Assert.AreEqual("'Sheet 1'!$A$1:$D$1", ws1.NamedRanges.First().RefersTo);
                    Assert.AreEqual("'Sheet 1'!A1:D1",
                        ws1.NamedRanges.First().Ranges.Single().RangeAddress.ToString(XLReferenceStyle.A1, true));

                    Assert.AreEqual(3, wb.NamedRanges.Count());

                    Assert.AreEqual("Named range 2", wb.NamedRanges.ElementAt(0).Name);
                    Assert.AreEqual(XLNamedRangeScope.Workbook, wb.NamedRanges.ElementAt(0).Scope);
                    Assert.AreEqual("'Sheet 1'!$A$2:$D$2", wb.NamedRanges.ElementAt(0).RefersTo);
                    Assert.AreEqual("'Sheet 1'!A2:D2",
                        wb.NamedRanges.ElementAt(0).Ranges.Single().RangeAddress.ToString(XLReferenceStyle.A1, true));

                    Assert.AreEqual("Named range 4", wb.NamedRanges.ElementAt(1).Name);
                    Assert.AreEqual(XLNamedRangeScope.Workbook, wb.NamedRanges.ElementAt(1).Scope);
                    Assert.AreEqual("#REF!", wb.NamedRanges.ElementAt(1).RefersTo);
                    Assert.IsFalse(wb.NamedRanges.ElementAt(1).Ranges.Any());

                    Assert.AreEqual("Named range 5", wb.NamedRanges.ElementAt(2).Name);
                    Assert.AreEqual(XLNamedRangeScope.Workbook, wb.NamedRanges.ElementAt(2).Scope);
                    Assert.AreEqual("'Sheet 1'!$A$5:$D$5,#REF!", wb.NamedRanges.ElementAt(2).RefersTo);
                    Assert.AreEqual(1, wb.NamedRanges.ElementAt(2).Ranges.Count);
                    Assert.AreEqual("'Sheet 1'!A5:D5",
                        wb.NamedRanges.ElementAt(2).Ranges.Single().RangeAddress.ToString(XLReferenceStyle.A1, true));
                }
            }
        }

        [Test]
        public void TestInvalidNamedRangeOnWorkbookScope()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.FirstCell().SetValue("Column1");
                ws.FirstCell().CellRight().SetValue("Column2").Style.Font.SetBold();
                ws.FirstCell().CellRight(2).SetValue("Column3");

                Assert.Throws<ArgumentException>(() => wb.NamedRanges.Add("MyRange", "A1:C1"));
            }
        }

        [Test]
        public void WbContainsWsNamedRange()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().AddToNamed("Name", XLScope.Worksheet);

            Assert.IsTrue(wb.NamedRanges.Contains("Sheet1!Name"));
            Assert.IsFalse(wb.NamedRanges.Contains("Sheet1!NameX"));

            Assert.IsNotNull(wb.NamedRange("Sheet1!Name"));
            Assert.IsNull(wb.NamedRange("Sheet1!NameX"));

            Boolean result1 = wb.NamedRanges.TryGetValue("Sheet1!Name", out IXLNamedRange range1);
            Assert.IsTrue(result1);
            Assert.IsNotNull(range1);
            Assert.AreEqual(XLNamedRangeScope.Worksheet, range1.Scope);

            Boolean result2 = wb.NamedRanges.TryGetValue("Sheet1!NameX", out IXLNamedRange range2);
            Assert.IsFalse(result2);
            Assert.IsNull(range2);
        }

        [Test]
        public void WorkbookContainsNamedRange()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().AddToNamed("Name");

            Assert.IsTrue(wb.NamedRanges.Contains("Name"));
            Assert.IsFalse(wb.NamedRanges.Contains("NameX"));

            Assert.IsNotNull(wb.NamedRange("Name"));
            Assert.IsNull(wb.NamedRange("NameX"));

            Boolean result1 = wb.NamedRanges.TryGetValue("Name", out IXLNamedRange range1);
            Assert.IsTrue(result1);
            Assert.IsNotNull(range1);

            Boolean result2 = wb.NamedRanges.TryGetValue("NameX", out IXLNamedRange range2);
            Assert.IsFalse(result2);
            Assert.IsNull(range2);
        }

        [Test]
        public void WorksheetContainsNamedRange()
        {
            IXLWorksheet ws = new XLWorkbook().AddWorksheet("Sheet1");
            ws.FirstCell().AddToNamed("Name", XLScope.Worksheet);

            Assert.IsTrue(ws.NamedRanges.Contains("Name"));
            Assert.IsFalse(ws.NamedRanges.Contains("NameX"));

            Assert.IsNotNull(ws.NamedRange("Name"));
            Assert.IsNull(ws.NamedRange("NameX"));

            Boolean result1 = ws.NamedRanges.TryGetValue("Name", out IXLNamedRange range1);
            Assert.IsTrue(result1);
            Assert.IsNotNull(range1);

            Boolean result2 = ws.NamedRanges.TryGetValue("NameX", out IXLNamedRange range2);
            Assert.IsFalse(result2);
            Assert.IsNull(range2);
        }

        [Test]
        public void NamedRangeWithSameNameAsAFunction()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            var a1 = ws.FirstCell();
            var a2 = a1.CellBelow();

            a1.SetValue(5).AddToNamed("RAND");
            a2.FormulaA1 = "=RAND * 10";

            Assert.AreEqual(50, a2.GetDouble());
        }
    }
}
