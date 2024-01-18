// Keep this file CodeMaid organised and cleaned
using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Parser;

namespace ClosedXML.Tests.Excel
{
    [TestFixture]
    public class NamedRangesTests
    {
        [Test]
        public void Formula_must_be_valid()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            Assert.Throws<ParsingException>(() => wb.DefinedNames.Add("Test", "SUM(Sheet7!A4"));
        }

        [Test]
        public void CanEvaluateNamedMultiRange()
        {
            using (var wb = new XLWorkbook())
            {
                var ws1 = wb.AddWorksheet("Sheet1");
                ws1.Range("A1:C1").Value = 1;
                ws1.Range("A3:C3").Value = 3;
                wb.DefinedNames.Add("TEST", ws1.Ranges("A1:C1,A3:C3"));

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

            Assert.AreEqual(1, wb.Cell("value1").Value);
            Assert.AreEqual(1, wb.Range("value1").FirstCell().Value);

            Assert.AreEqual(1, ws1.Cell("value1").Value);
            Assert.AreEqual(1, ws1.Range("value1").FirstCell().Value);

            var ws2 = wb.Worksheets.Add("Sheet2");

            ws2.Cell("A1").SetFormulaA1("=value1").AddToNamed("value2");

            Assert.AreEqual(1, wb.Cell("value2").Value);
            Assert.AreEqual(1, wb.Range("value2").FirstCell().Value);

            Assert.AreEqual(1, ws2.Cell("value1").Value);
            Assert.AreEqual(1, ws2.Range("value1").FirstCell().Value);

            Assert.AreEqual(1, ws2.Cell("value2").Value);
            Assert.AreEqual(1, ws2.Range("value2").FirstCell().Value);
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
                wb.DefinedNames.Add("Named range 5", new XLRanges
                {
                    ws1.Range("A5:D5"),
                    ws3.Range("A5:D5")
                });

                ws2.Delete();
                ws3.Delete();

                var globalValidRanges = wb.DefinedNames.ValidNamedRanges();
                var globalInvalidRanges = wb.DefinedNames.InvalidNamedRanges();
                var localValidRanges = ws1.DefinedNames.ValidNamedRanges();
                var localInvalidRanges = ws1.DefinedNames.InvalidNamedRanges();

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
                var dn1 = wb.DefinedNames.Add("TEST", "=0.1");

                Assert.IsTrue(wb.DefinedNames.TryGetValue("TEST", out _));
                Assert.IsFalse(wb.DefinedNames.TryGetValue("TEST1", out _));

                dn1.Name = "TEST1";

                Assert.IsFalse(wb.DefinedNames.TryGetValue("TEST", out _));
                Assert.IsTrue(wb.DefinedNames.TryGetValue("TEST1", out _));

                var dn2 = wb.DefinedNames.Add("TEST2", "=TEST1*2");

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
        public void Can_save_and_load_defined_names()
        {
            using (var ms = new MemoryStream())
            {
                using (var wb = new XLWorkbook())
                {
                    var sheet1 = wb.Worksheets.Add("Sheet1");
                    var sheet2 = wb.Worksheets.Add("Sheet2");

                    wb.DefinedNames.Add("wbNamedRange",
                        "Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G");
                    sheet1.DefinedNames.Add("sheet1NamedRange",
                        "Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G");
                    sheet2.DefinedNames.Add("sheet2NamedRange", "Sheet1!A1,Sheet2!A1");

                    wb.SaveAs(ms);
                }

                using (var wb = new XLWorkbook(ms))
                {
                    var sheet1 = wb.Worksheet("Sheet1");
                    var sheet2 = wb.Worksheet("Sheet2");

                    Assert.AreEqual(1, wb.DefinedNames.Count());
                    Assert.AreEqual("wbNamedRange", wb.DefinedNames.Single().Name);
                    Assert.AreEqual("Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G", wb.DefinedNames.Single().RefersTo);
                    Assert.AreEqual(5, wb.DefinedNames.Single().Ranges.Count);

                    Assert.AreEqual(1, sheet1.DefinedNames.Count());
                    Assert.AreEqual("sheet1NamedRange", sheet1.DefinedNames.Single().Name);
                    Assert.AreEqual("Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G", sheet1.DefinedNames.Single().RefersTo);
                    Assert.AreEqual(5, sheet1.DefinedNames.Single().Ranges.Count);

                    Assert.AreEqual(1, sheet2.DefinedNames.Count());
                    Assert.AreEqual("sheet2NamedRange", sheet2.DefinedNames.Single().Name);
                    Assert.AreEqual("Sheet1!A1,Sheet2!A1", sheet2.DefinedNames.Single().RefersTo);
                    Assert.AreEqual(2, sheet2.DefinedNames.Single().Ranges.Count);
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
            var original = ws1.DefinedNames.Add("Named range", ranges);

            var copy = original.CopyTo(ws2);

            Assert.AreEqual(1, ws1.DefinedNames.Count());
            Assert.AreEqual(1, ws2.DefinedNames.Count());
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
        public void Copy_table_references_to_different_worksheet()
        {
            // When sheet-scoped name references a table and there is a table with same area in the
            // copied sheet, the copied defined name changes table reference to a new table. If
            // range differs, table reference is not modified.
            using var wb = new XLWorkbook();
            var orgSheet = wb.AddWorksheet();
            orgSheet.Cell("A1").InsertTable(new[] { "Data", "A", "B" }, "OrgTable", true);
            orgSheet.Cell("C1").InsertTable(new[] { "Data", "A", "B" }, "MiscTable", true);
            var originalName = orgSheet.DefinedNames.Add("TableName", "SUM(OrgTable[Data], MiscTable[Data])");

            var copySheet = wb.AddWorksheet();
            copySheet.Cell("A1").InsertTable(new[] { "Data", "A", "B" }, "CopyTable", true);

            originalName.CopyTo(copySheet);

            var copyName = copySheet.DefinedNames.Single();
            Assert.AreEqual("TableName", copyName.Name);
            Assert.AreEqual("SUM(CopyTable[Data], MiscTable[Data])", copyName.RefersTo);
        }

        [Test]
        public void Copy_workbook_scoped_defined()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");
            var name = wb.DefinedNames.Add("Name", "Sheet!$A$1");

            var copySheet = wb.AddWorksheet();
            var ex = Assert.Throws<InvalidOperationException>(() => name.CopyTo(copySheet))!;
            Assert.AreEqual("Cannot copy workbook scoped defined name.", ex.Message);
        }

        [Test]
        public void Copy_defined_name_to_same_sheet()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("Sheet1");
            ws1.Range("B2:E6").AddToNamed("Named range", XLScope.Worksheet);
            var dn = ws1.DefinedName("Named range");

            TestDelegate action = () => dn.CopyTo(ws1);

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
                ws.DefinedNames.Add("MyRange", "A1:C1");

                ws.Column(1).Delete();

                Assert.IsTrue(ws.Cell("A1").Style.Font.Bold);
                Assert.AreEqual("Column3", ws.Cell("B1").Value);
                Assert.AreEqual(Blank.Value, ws.Cell("C1").Value);
            }
        }

        [Test]
        public void Formula_is_updated_on_sheet_rename()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Old name");
            var bookScopedName = wb.DefinedNames.Add("TEST", "ABS('Old name'!$B$5)");
            var sheetScopedName = ws.DefinedNames.Add("TEST1", "'Old name'!$D$7:$F$14");

            ws.Name = "Renamed";

            Assert.AreEqual("ABS(Renamed!$B$5)", bookScopedName.RefersTo);
            Assert.AreEqual("Renamed!$B$5:$B$5", bookScopedName.Ranges.ToString());

            Assert.AreEqual("Renamed!$D$7:$F$14", sheetScopedName.RefersTo);
            Assert.AreEqual("Renamed!$D$7:$F$14", sheetScopedName.Ranges.ToString());
        }

        [Test]
        public void MovingRanges()
        {
            var wb = new XLWorkbook();

            IXLWorksheet sheet1 = wb.Worksheets.Add("Sheet1");
            IXLWorksheet sheet2 = wb.Worksheets.Add("Sheet2");

            wb.DefinedNames.Add("wbNamedRange",
                "Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G");
            sheet1.DefinedNames.Add("sheet1NamedRange",
                "Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G");
            sheet2.DefinedNames.Add("sheet2NamedRange", "Sheet1!A1,Sheet2!A1");

            sheet1.Row(1).InsertRowsAbove(2);
            sheet1.Row(1).Delete();
            sheet1.Column(1).InsertColumnsBefore(2);
            sheet1.Column(1).Delete();

            Assert.AreEqual("Sheet1!$C$3,Sheet1!$C$4:$D$4,Sheet2!$D$3:$D$4,Sheet1!$7:$8,Sheet1!$G:$H",
                wb.DefinedNames.First().RefersTo);
            Assert.AreEqual("Sheet1!$C$3,Sheet1!$C$4:$D$4,Sheet2!$D$3:$D$4,Sheet1!$7:$8,Sheet1!$G:$H",
                sheet1.DefinedNames.First().RefersTo);
            Assert.AreEqual("Sheet1!B2,Sheet2!A1", sheet2.DefinedNames.First().RefersTo);

            wb.DefinedNames.ForEach(dn => Assert.AreEqual(XLNamedRangeScope.Workbook, dn.Scope));
            sheet1.DefinedNames.ForEach(dn => Assert.AreEqual(XLNamedRangeScope.Worksheet, dn.Scope));
            sheet2.DefinedNames.ForEach(dn => Assert.AreEqual(XLNamedRangeScope.Worksheet, dn.Scope));
        }

        [Test, Ignore("Muted until shifting is fixed (see #880)")]
        public void NamedRangeBecomesInvalidOnRangeAndWorksheetDeleting()
        {
            using (var wb = new XLWorkbook())
            {
                var ws1 = wb.Worksheets.Add("Sheet 1");
                var ws2 = wb.Worksheets.Add("Sheet 2");
                ws1.Range("A1:B2").AddToNamed("Simple", XLScope.Workbook);
                wb.DefinedNames.Add("Compound", new XLRanges
                {
                    ws1.Range("C1:D2"),
                    ws2.Range("A10:D15")
                });

                ws1.Rows(1, 5).Delete();
                ws1.Delete();

                Assert.AreEqual(2, wb.DefinedNames.Count());
                Assert.AreEqual(0, wb.DefinedNames.ValidNamedRanges().Count());
                Assert.AreEqual("#REF!#REF!", wb.DefinedNames.ElementAt(0).RefersTo);
                Assert.AreEqual("#REF!#REF!,'Sheet 2'!A10:D15", wb.DefinedNames.ElementAt(0).RefersTo);
            }
        }

        [Test, Ignore("Muted until shifting is fixed (see #880)")]
        public void NamedRangeBecomesInvalidOnRangeDeleting()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet 1");
                ws.Range("A1:B2").AddToNamed("Simple", XLScope.Workbook);
                wb.DefinedNames.Add("Compound", new XLRanges
                {
                    ws.Range("C1:D2"),
                    ws.Range("A10:D15")
                });

                ws.Rows(1, 5).Delete();

                Assert.AreEqual(2, wb.DefinedNames.Count());
                Assert.AreEqual(0, wb.DefinedNames.ValidNamedRanges().Count());
                Assert.AreEqual("'Sheet 1'!#REF!", wb.DefinedNames.ElementAt(0).RefersTo);
                Assert.AreEqual("'Sheet 1'!#REF!,'Sheet 1'!A5:D10", wb.DefinedNames.ElementAt(0).RefersTo);
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
                    wb.DefinedNames.Add("TEST", "=0.1");
                    wb.DefinedNames.Add("TEST2", "=TEST*2");

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

                    wb.DefinedNames.Add("Multirange named range", new XLRanges
                    {
                        ws.Range("A5:D5"),
                        ws.Range("A15:D15")
                    });

                    wb.SaveAs(ms);
                }

                using (var wb = new XLWorkbook(ms))
                {
                    Assert.AreEqual(1, wb.DefinedNames.Count());
                    var nr = (XLDefinedName)wb.DefinedNames.Single();
                    Assert.AreEqual("'Sheet 1'!$A$5:$D$5,'Sheet 1'!$A$15:$D$15", nr.RefersTo);
                    Assert.AreEqual(2, nr.Ranges.Count);
                    Assert.AreEqual("'Sheet 1'!A5:D5", nr.Ranges.First().RangeAddress.ToString(XLReferenceStyle.A1, true));
                    Assert.AreEqual("'Sheet 1'!A15:D15", nr.Ranges.Last().RangeAddress.ToString(XLReferenceStyle.A1, true));
                    Assert.AreEqual(2, nr.SheetReferencesList.Count);
                    Assert.AreEqual("'Sheet 1'!$A$5:$D$5", nr.SheetReferencesList.First());
                    Assert.AreEqual("'Sheet 1'!$A$15:$D$15", nr.SheetReferencesList.Last());
                }
            }
        }

        [Test]
        public void Defined_names_referencing_sheet_range_become_invalid_when_sheet_is_deleted()
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
                wb.DefinedNames.Add("Named range 5", new XLRanges
                {
                    ws1.Range("A5:D5"),
                    ws3.Range("A5:D5")
                });

                ws2.Delete();
                ws3.Delete();

                Assert.AreEqual(1, ws1.DefinedNames.Count());
                Assert.AreEqual("Named range 1", ws1.DefinedNames.First().Name);
                Assert.AreEqual(XLNamedRangeScope.Worksheet, ws1.DefinedNames.First().Scope);
                Assert.AreEqual("'Sheet 1'!$A$1:$D$1", ws1.DefinedNames.First().RefersTo);
                Assert.AreEqual("'Sheet 1'!A1:D1", ws1.DefinedNames.First().Ranges.Single().RangeAddress.ToString(XLReferenceStyle.A1, true));

                Assert.AreEqual(3, wb.DefinedNames.Count());

                Assert.AreEqual("Named range 2", wb.DefinedNames.ElementAt(0).Name);
                Assert.AreEqual(XLNamedRangeScope.Workbook, wb.DefinedNames.ElementAt(0).Scope);
                Assert.AreEqual("'Sheet 1'!$A$2:$D$2", wb.DefinedNames.ElementAt(0).RefersTo);
                Assert.AreEqual("'Sheet 1'!A2:D2", wb.DefinedNames.ElementAt(0).Ranges.Single().RangeAddress.ToString(XLReferenceStyle.A1, true));

                Assert.AreEqual("Named range 4", wb.DefinedNames.ElementAt(1).Name);
                Assert.AreEqual(XLNamedRangeScope.Workbook, wb.DefinedNames.ElementAt(1).Scope);
                Assert.AreEqual("#REF!", wb.DefinedNames.ElementAt(1).RefersTo);
                Assert.IsFalse(wb.DefinedNames.ElementAt(1).Ranges.Any());

                Assert.AreEqual("Named range 5", wb.DefinedNames.ElementAt(2).Name);
                Assert.AreEqual(XLNamedRangeScope.Workbook, wb.DefinedNames.ElementAt(2).Scope);
                Assert.AreEqual("'Sheet 1'!$A$5:$D$5,#REF!", wb.DefinedNames.ElementAt(2).RefersTo);
                Assert.AreEqual(1, wb.DefinedNames.ElementAt(2).Ranges.Count);
                Assert.AreEqual("'Sheet 1'!A5:D5", wb.DefinedNames.ElementAt(2).Ranges.Single().RangeAddress.ToString(XLReferenceStyle.A1, true));
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
                    Assert.AreEqual("#REF!", wb.DefinedNames.Single().RefersTo);
                }
            }
        }

        [Test]
        public void Only_worksheet_scoped_defined_names_are_copied_when_sheet_is_copied()
        {
            using (var wb = new XLWorkbook())
            {
                var ws1 = wb.AddWorksheet("Sheet1");
                ws1.FirstCell().InsertData(Enumerable.Range(1, 10));
                wb.DefinedNames.Add("wbNamedRange", ws1.Range("A1:A10"));
                ws1.DefinedNames.Add("wsNamedRange", ws1.Range("A3"));

                var ws2 = wb.AddWorksheet("Sheet2");
                ws2.FirstCell().InsertData(Enumerable.Range(101, 10));
                ws1.DefinedNames.Add("wsNamedRangeAcrossSheets", ws2.Range("A4"));

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
                    wb.DefinedName("wbNamedRange").Ranges.First().RangeAddress.ToStringRelative(true));
                Assert.AreEqual("Copy!A3:A3",
                    wsCopy.DefinedName("wsNamedRange").Ranges.First().RangeAddress.ToStringRelative(true));
                Assert.AreEqual("Sheet2!A4:A4",
                    wsCopy.DefinedName("wsNamedRangeAcrossSheets").Ranges.First().RangeAddress.ToStringRelative(true));
            }
        }

        [Test]
        public void Saved_defined_names_become_invalid_on_sheet_deleting()
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
                    wb.DefinedNames.Add("Named range 5", new XLRanges
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
                    Assert.AreEqual(1, ws1.DefinedNames.Count());
                    Assert.AreEqual("Named range 1", ws1.DefinedNames.First().Name);
                    Assert.AreEqual(XLNamedRangeScope.Worksheet, ws1.DefinedNames.First().Scope);
                    Assert.AreEqual("'Sheet 1'!$A$1:$D$1", ws1.DefinedNames.First().RefersTo);
                    Assert.AreEqual("'Sheet 1'!A1:D1",
                        ws1.DefinedNames.First().Ranges.Single().RangeAddress.ToString(XLReferenceStyle.A1, true));

                    Assert.AreEqual(3, wb.DefinedNames.Count());

                    Assert.AreEqual("Named range 2", wb.DefinedNames.ElementAt(0).Name);
                    Assert.AreEqual(XLNamedRangeScope.Workbook, wb.DefinedNames.ElementAt(0).Scope);
                    Assert.AreEqual("'Sheet 1'!$A$2:$D$2", wb.DefinedNames.ElementAt(0).RefersTo);
                    Assert.AreEqual("'Sheet 1'!A2:D2",
                        wb.DefinedNames.ElementAt(0).Ranges.Single().RangeAddress.ToString(XLReferenceStyle.A1, true));

                    Assert.AreEqual("Named range 4", wb.DefinedNames.ElementAt(1).Name);
                    Assert.AreEqual(XLNamedRangeScope.Workbook, wb.DefinedNames.ElementAt(1).Scope);
                    Assert.AreEqual("#REF!", wb.DefinedNames.ElementAt(1).RefersTo);
                    Assert.IsFalse(wb.DefinedNames.ElementAt(1).Ranges.Any());

                    Assert.AreEqual("Named range 5", wb.DefinedNames.ElementAt(2).Name);
                    Assert.AreEqual(XLNamedRangeScope.Workbook, wb.DefinedNames.ElementAt(2).Scope);
                    Assert.AreEqual("'Sheet 1'!$A$5:$D$5,#REF!", wb.DefinedNames.ElementAt(2).RefersTo);
                    Assert.AreEqual(1, wb.DefinedNames.ElementAt(2).Ranges.Count);
                    Assert.AreEqual("'Sheet 1'!A5:D5",
                        wb.DefinedNames.ElementAt(2).Ranges.Single().RangeAddress.ToString(XLReferenceStyle.A1, true));
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

                Assert.Throws<ArgumentException>(() => wb.DefinedNames.Add("MyRange", "A1:C1"));
            }
        }

        [Test]
        public void WbContainsWsNamedRange()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().AddToNamed("Name", XLScope.Worksheet);

            Assert.IsTrue(wb.DefinedNames.Contains("Sheet1!Name"));
            Assert.IsFalse(wb.DefinedNames.Contains("Sheet1!NameX"));

            Assert.IsNotNull(wb.DefinedName("Sheet1!Name"));
            Assert.IsNull(wb.DefinedName("Sheet1!NameX"));

            Boolean found1 = wb.DefinedNames.TryGetValue("Sheet1!Name", out var definedName1);
            Assert.IsTrue(found1);
            Assert.IsNotNull(definedName1);
            Assert.AreEqual(XLNamedRangeScope.Worksheet, definedName1.Scope);

            Boolean found2 = wb.DefinedNames.TryGetValue("Sheet1!NameX", out var definedName2);
            Assert.IsFalse(found2);
            Assert.IsNull(definedName2);
        }

        [Test]
        public void WorkbookContainsNamedRange()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().AddToNamed("Name");

            Assert.IsTrue(wb.DefinedNames.Contains("Name"));
            Assert.IsFalse(wb.DefinedNames.Contains("NameX"));

            Assert.IsNotNull(wb.DefinedName("Name"));
            Assert.IsNull(wb.DefinedName("NameX"));

            Boolean found1 = wb.DefinedNames.TryGetValue("Name", out var definedName1);
            Assert.IsTrue(found1);
            Assert.IsNotNull(definedName1);

            Boolean found2 = wb.DefinedNames.TryGetValue("NameX", out var definedName2);
            Assert.IsFalse(found2);
            Assert.IsNull(definedName2);
        }

        [Test]
        public void WorksheetContainsNamedRange()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().AddToNamed("Name", XLScope.Worksheet);

            Assert.IsTrue(ws.DefinedNames.Contains("Name"));
            Assert.IsFalse(ws.DefinedNames.Contains("NameX"));

            Assert.IsNotNull(ws.DefinedName("Name"));
            Assert.Throws<KeyNotFoundException>(() => ws.DefinedName("NameX"));

            Boolean found1 = ws.DefinedNames.TryGetValue("Name", out var definedName1);
            Assert.IsTrue(found1);
            Assert.IsNotNull(definedName1);

            Boolean found2 = ws.DefinedNames.TryGetValue("NameX", out var definedName2);
            Assert.IsFalse(found2);
            Assert.IsNull(definedName2);
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
