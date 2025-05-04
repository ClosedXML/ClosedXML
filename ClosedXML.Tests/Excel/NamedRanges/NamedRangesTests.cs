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
            using var wb = new XLWorkbook();
            var ws1 = wb.AddWorksheet("Sheet1");
            ws1.Range("A1:C1").Value = 1;
            ws1.Range("A3:C3").Value = 3;
            wb.DefinedNames.Add("TEST", ws1.Ranges("A1:C1,A3:C3"));

            ws1.Cell(2, 1).FormulaA1 = "=SUM(TEST)";

            Assert.That((double)ws1.Cell(2, 1).Value, Is.EqualTo(12.0).Within(XLHelper.Epsilon));
        }

        [Test]
        public void CanGetNamedFromAnother()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("Sheet1");
            ws1.Cell("A1").SetValue(1).AddToNamed("value1");

            Assert.Multiple(() =>
            {
                Assert.That(wb.Cell("value1").Value, Is.EqualTo(1));
                Assert.That(wb.Range("value1").FirstCell().Value, Is.EqualTo(1));

                Assert.That(ws1.Cell("value1").Value, Is.EqualTo(1));
                Assert.That(ws1.Range("value1").FirstCell().Value, Is.EqualTo(1));
            });

            var ws2 = wb.Worksheets.Add("Sheet2");

            ws2.Cell("A1").SetFormulaA1("=value1").AddToNamed("value2");

            Assert.Multiple(() =>
            {
                Assert.That(wb.Cell("value2").Value, Is.EqualTo(1));
                Assert.That(wb.Range("value2").FirstCell().Value, Is.EqualTo(1));

                Assert.That(ws2.Cell("value1").Value, Is.EqualTo(1));
                Assert.That(ws2.Range("value1").FirstCell().Value, Is.EqualTo(1));

                Assert.That(ws2.Cell("value2").Value, Is.EqualTo(1));
                Assert.That(ws2.Range("value2").FirstCell().Value, Is.EqualTo(1));
            });
        }

        [Test]
        public void CanGetValidNamedRanges()
        {
            using var wb = new XLWorkbook();
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

            Assert.Multiple(() =>
            {
                Assert.That(globalValidRanges.Count(), Is.EqualTo(1));
                Assert.That(globalValidRanges.First().Name, Is.EqualTo("Named range 2"));

                Assert.That(globalInvalidRanges.Count(), Is.EqualTo(2));
                Assert.That(globalInvalidRanges.First().Name, Is.EqualTo("Named range 4"));
                Assert.That(globalInvalidRanges.Last().Name, Is.EqualTo("Named range 5"));

                Assert.That(localValidRanges.Count(), Is.EqualTo(1));
                Assert.That(localValidRanges.First().Name, Is.EqualTo("Named range 1"));

                Assert.That(localInvalidRanges.Count(), Is.EqualTo(0));
            });
        }

        [Test]
        public void CanRenameNamedRange()
        {
            using var wb = new XLWorkbook();
            var ws1 = wb.AddWorksheet("Sheet1");
            var dn1 = wb.DefinedNames.Add("TEST", "=0.1");

            Assert.Multiple(() =>
            {
                Assert.That(wb.DefinedNames.TryGetValue("TEST", out _), Is.True);
                Assert.That(wb.DefinedNames.TryGetValue("TEST1", out _), Is.False);
            });

            dn1.Name = "TEST1";

            Assert.Multiple(() =>
            {
                Assert.That(wb.DefinedNames.TryGetValue("TEST", out _), Is.False);
                Assert.That(wb.DefinedNames.TryGetValue("TEST1", out _), Is.True);
            });

            var dn2 = wb.DefinedNames.Add("TEST2", "=TEST1*2");

            ws1.Cell(1, 1).FormulaA1 = "TEST1";
            ws1.Cell(2, 1).FormulaA1 = "TEST1*10";
            ws1.Cell(3, 1).FormulaA1 = "TEST2";
            ws1.Cell(4, 1).FormulaA1 = "TEST2*3";

            Assert.Multiple(() =>
            {
                Assert.That((double)ws1.Cell(1, 1).Value, Is.EqualTo(0.1).Within(XLHelper.Epsilon));
                Assert.That((double)ws1.Cell(2, 1).Value, Is.EqualTo(1.0).Within(XLHelper.Epsilon));
                Assert.That((double)ws1.Cell(3, 1).Value, Is.EqualTo(0.2).Within(XLHelper.Epsilon));
                Assert.That((double)ws1.Cell(4, 1).Value, Is.EqualTo(0.6).Within(XLHelper.Epsilon));
            });
        }

        [Test]
        public void Can_save_and_load_defined_names()
        {
            using var ms = new MemoryStream();
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

                Assert.Multiple(() =>
                {
                    Assert.That(wb.DefinedNames.Count(), Is.EqualTo(1));
                    Assert.That(wb.DefinedNames.Single().Name, Is.EqualTo("wbNamedRange"));
                    Assert.That(wb.DefinedNames.Single().RefersTo, Is.EqualTo("Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G"));
                    Assert.That(wb.DefinedNames.Single().Ranges, Has.Count.EqualTo(5));

                    Assert.That(sheet1.DefinedNames.Count(), Is.EqualTo(1));
                    Assert.That(sheet1.DefinedNames.Single().Name, Is.EqualTo("sheet1NamedRange"));
                    Assert.That(sheet1.DefinedNames.Single().RefersTo, Is.EqualTo("Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G"));
                    Assert.That(sheet1.DefinedNames.Single().Ranges, Has.Count.EqualTo(5));

                    Assert.That(sheet2.DefinedNames.Count(), Is.EqualTo(1));
                    Assert.That(sheet2.DefinedNames.Single().Name, Is.EqualTo("sheet2NamedRange"));
                    Assert.That(sheet2.DefinedNames.Single().RefersTo, Is.EqualTo("Sheet1!A1,Sheet2!A1"));
                    Assert.That(sheet2.DefinedNames.Single().Ranges, Has.Count.EqualTo(2));
                });
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

            Assert.Multiple(() =>
            {
                Assert.That(ws1.DefinedNames.Count(), Is.EqualTo(1));
                Assert.That(ws2.DefinedNames.Count(), Is.EqualTo(1));
                Assert.That(original.Ranges, Has.Count.EqualTo(2));
                Assert.That(copy.Ranges, Has.Count.EqualTo(2));
                Assert.That(copy.Name, Is.EqualTo(original.Name));
                Assert.That(copy.Scope, Is.EqualTo(original.Scope));
            });
            Assert.That(original.Ranges.First().RangeAddress.ToString(XLReferenceStyle.A1, true), Is.EqualTo("Sheet1!B2:E6"));
            Assert.That(original.Ranges.Last().RangeAddress.ToString(XLReferenceStyle.A1, true), Is.EqualTo("Sheet2!D1:E2"));
            Assert.That(copy.Ranges.First().RangeAddress.ToString(XLReferenceStyle.A1, true), Is.EqualTo("Sheet2!D1:E2"));
            Assert.That(copy.Ranges.Last().RangeAddress.ToString(XLReferenceStyle.A1, true), Is.EqualTo("Sheet2!B2:E6"));
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
            Assert.Multiple(() =>
            {
                Assert.That(copyName.Name, Is.EqualTo("TableName"));
                Assert.That(copyName.RefersTo, Is.EqualTo("SUM(CopyTable[Data], MiscTable[Data])"));
            });
        }

        [Test]
        public void Copy_workbook_scoped_defined()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");
            var name = wb.DefinedNames.Add("Name", "Sheet!$A$1");

            var copySheet = wb.AddWorksheet();
            var ex = Assert.Throws<InvalidOperationException>(() => name.CopyTo(copySheet))!;
            Assert.That(ex.Message, Is.EqualTo("Cannot copy workbook scoped defined name."));
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
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Column1");
            ws.FirstCell().CellRight().SetValue("Column2").Style.Font.SetBold();
            ws.FirstCell().CellRight(2).SetValue("Column3");
            ws.DefinedNames.Add("MyRange", "A1:C1");

            ws.Column(1).Delete();

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell("A1").Style.Font.Bold, Is.True);
                Assert.That(ws.Cell("B1").Value, Is.EqualTo("Column3"));
                Assert.That(ws.Cell("C1").Value, Is.EqualTo(Blank.Value));
            });
        }

        [Test]
        public void Formula_is_updated_on_sheet_rename()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Old name");
            var bookScopedName = wb.DefinedNames.Add("TEST", "ABS('Old name'!$B$5)");
            var sheetScopedName = ws.DefinedNames.Add("TEST1", "'Old name'!$D$7:$F$14");

            ws.Name = "Renamed";

            Assert.Multiple(() =>
            {
                Assert.That(bookScopedName.RefersTo, Is.EqualTo("ABS(Renamed!$B$5)"));
                Assert.That(bookScopedName.Ranges.ToString(), Is.EqualTo("Renamed!$B$5:$B$5"));

                Assert.That(sheetScopedName.RefersTo, Is.EqualTo("Renamed!$D$7:$F$14"));
                Assert.That(sheetScopedName.Ranges.ToString(), Is.EqualTo("Renamed!$D$7:$F$14"));
            });
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

            Assert.Multiple(() =>
            {
                Assert.That(wb.DefinedNames.First().RefersTo,
                            Is.EqualTo("Sheet1!$C$3,Sheet1!$C$4:$D$4,Sheet2!$D$3:$D$4,Sheet1!$7:$8,Sheet1!$G:$H"));
                Assert.That(sheet1.DefinedNames.First().RefersTo,
                    Is.EqualTo("Sheet1!$C$3,Sheet1!$C$4:$D$4,Sheet2!$D$3:$D$4,Sheet1!$7:$8,Sheet1!$G:$H"));
                Assert.That(sheet2.DefinedNames.First().RefersTo, Is.EqualTo("Sheet1!B2,Sheet2!A1"));
            });

            wb.DefinedNames.ForEach(dn => Assert.That(dn.Scope, Is.EqualTo(XLNamedRangeScope.Workbook)));
            sheet1.DefinedNames.ForEach(dn => Assert.That(dn.Scope, Is.EqualTo(XLNamedRangeScope.Worksheet)));
            sheet2.DefinedNames.ForEach(dn => Assert.That(dn.Scope, Is.EqualTo(XLNamedRangeScope.Worksheet)));
        }

        [Test, Ignore("Muted until shifting is fixed (see #880)")]
        public void NamedRangeBecomesInvalidOnRangeAndWorksheetDeleting()
        {
            using var wb = new XLWorkbook();
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

            Assert.Multiple(() =>
            {
                Assert.That(wb.DefinedNames.Count(), Is.EqualTo(2));
                Assert.That(wb.DefinedNames.ValidNamedRanges().Count(), Is.EqualTo(0));
                Assert.That(wb.DefinedNames.ElementAt(0).RefersTo, Is.EqualTo("#REF!#REF!"));
            });
            Assert.That(wb.DefinedNames.ElementAt(0).RefersTo, Is.EqualTo("#REF!#REF!,'Sheet 2'!A10:D15"));
        }

        [Test, Ignore("Muted until shifting is fixed (see #880)")]
        public void NamedRangeBecomesInvalidOnRangeDeleting()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet 1");
            ws.Range("A1:B2").AddToNamed("Simple", XLScope.Workbook);
            wb.DefinedNames.Add("Compound", new XLRanges
            {
                ws.Range("C1:D2"),
                ws.Range("A10:D15")
            });

            ws.Rows(1, 5).Delete();

            Assert.Multiple(() =>
            {
                Assert.That(wb.DefinedNames.Count(), Is.EqualTo(2));
                Assert.That(wb.DefinedNames.ValidNamedRanges().Count(), Is.EqualTo(0));
                Assert.That(wb.DefinedNames.ElementAt(0).RefersTo, Is.EqualTo("'Sheet 1'!#REF!"));
            });
            Assert.That(wb.DefinedNames.ElementAt(0).RefersTo, Is.EqualTo("'Sheet 1'!#REF!,'Sheet 1'!A5:D10"));
        }

        [Test]
        public void NamedRangeMayReferToExpression()
        {
            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                var ws1 = wb.AddWorksheet("Sheet1");
                wb.DefinedNames.Add("TEST", "=0.1");
                wb.DefinedNames.Add("TEST2", "=TEST*2");

                ws1.Cell(1, 1).FormulaA1 = "TEST";
                ws1.Cell(2, 1).FormulaA1 = "TEST*10";
                ws1.Cell(3, 1).FormulaA1 = "TEST2";
                ws1.Cell(4, 1).FormulaA1 = "TEST2*3";

                Assert.Multiple(() =>
                {
                    Assert.That((double)ws1.Cell(1, 1).Value, Is.EqualTo(0.1).Within(XLHelper.Epsilon));
                    Assert.That((double)ws1.Cell(2, 1).Value, Is.EqualTo(1.0).Within(XLHelper.Epsilon));
                    Assert.That((double)ws1.Cell(3, 1).Value, Is.EqualTo(0.2).Within(XLHelper.Epsilon));
                    Assert.That((double)ws1.Cell(4, 1).Value, Is.EqualTo(0.6).Within(XLHelper.Epsilon));
                });

                wb.SaveAs(ms);
            }

            using (var wb = new XLWorkbook(ms))
            {
                var ws1 = wb.Worksheets.First();

                Assert.Multiple(() =>
                {
                    Assert.That((double)ws1.Cell(1, 1).Value, Is.EqualTo(0.1).Within(XLHelper.Epsilon));
                    Assert.That((double)ws1.Cell(2, 1).Value, Is.EqualTo(1.0).Within(XLHelper.Epsilon));
                    Assert.That((double)ws1.Cell(3, 1).Value, Is.EqualTo(0.2).Within(XLHelper.Epsilon));
                    Assert.That((double)ws1.Cell(4, 1).Value, Is.EqualTo(0.6).Within(XLHelper.Epsilon));
                });
            }
        }

        [Test]
        public void NamedRangeReferringToMultipleRangesCanBeSavedAndLoaded()
        {
            using var ms = new MemoryStream();
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
                Assert.That(wb.DefinedNames.Count(), Is.EqualTo(1));
                var nr = (XLDefinedName)wb.DefinedNames.Single();
                Assert.Multiple(() =>
                {
                    Assert.That(nr.RefersTo, Is.EqualTo("'Sheet 1'!$A$5:$D$5,'Sheet 1'!$A$15:$D$15"));
                    Assert.That(nr.Ranges, Has.Count.EqualTo(2));
                });
                Assert.That(nr.Ranges.First().RangeAddress.ToString(XLReferenceStyle.A1, true), Is.EqualTo("'Sheet 1'!A5:D5"));
                Assert.That(nr.Ranges.Last().RangeAddress.ToString(XLReferenceStyle.A1, true), Is.EqualTo("'Sheet 1'!A15:D15"));
                Assert.That(nr.SheetReferencesList, Has.Count.EqualTo(2));
                Assert.Multiple(() =>
                {
                    Assert.That(nr.SheetReferencesList.First(), Is.EqualTo("'Sheet 1'!$A$5:$D$5"));
                    Assert.That(nr.SheetReferencesList.Last(), Is.EqualTo("'Sheet 1'!$A$15:$D$15"));
                });
            }
        }

        [Test]
        public void Defined_names_referencing_sheet_range_become_invalid_when_sheet_is_deleted()
        {
            using var wb = new XLWorkbook();
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

            Assert.Multiple(() =>
            {
                Assert.That(ws1.DefinedNames.Count(), Is.EqualTo(1));
                Assert.That(ws1.DefinedNames.First().Name, Is.EqualTo("Named range 1"));
                Assert.That(ws1.DefinedNames.First().Scope, Is.EqualTo(XLNamedRangeScope.Worksheet));
                Assert.That(ws1.DefinedNames.First().RefersTo, Is.EqualTo("'Sheet 1'!$A$1:$D$1"));
                Assert.That(ws1.DefinedNames.First().Ranges.Single().RangeAddress.ToString(XLReferenceStyle.A1, true), Is.EqualTo("'Sheet 1'!A1:D1"));
            });

            Assert.That(wb.DefinedNames.Count(), Is.EqualTo(3));

            Assert.That(wb.DefinedNames.ElementAt(0).Name, Is.EqualTo("Named range 2"));
            Assert.Multiple(() =>
            {
                Assert.That(wb.DefinedNames.ElementAt(0).Scope, Is.EqualTo(XLNamedRangeScope.Workbook));
                Assert.That(wb.DefinedNames.ElementAt(0).RefersTo, Is.EqualTo("'Sheet 1'!$A$2:$D$2"));
                Assert.That(wb.DefinedNames.ElementAt(0).Ranges.Single().RangeAddress.ToString(XLReferenceStyle.A1, true), Is.EqualTo("'Sheet 1'!A2:D2"));
            });

            Assert.That(wb.DefinedNames.ElementAt(1).Name, Is.EqualTo("Named range 4"));
            Assert.That(wb.DefinedNames.ElementAt(1).Scope, Is.EqualTo(XLNamedRangeScope.Workbook));
            Assert.Multiple(() =>
            {
                Assert.That(wb.DefinedNames.ElementAt(1).RefersTo, Is.EqualTo("#REF!"));
                Assert.That(wb.DefinedNames.ElementAt(1).Ranges.Any(), Is.False);

                Assert.That(wb.DefinedNames.ElementAt(2).Name, Is.EqualTo("Named range 5"));
                Assert.That(wb.DefinedNames.ElementAt(2).Scope, Is.EqualTo(XLNamedRangeScope.Workbook));
                Assert.That(wb.DefinedNames.ElementAt(2).RefersTo, Is.EqualTo("'Sheet 1'!$A$5:$D$5,#REF!"));
                Assert.That(wb.DefinedNames.ElementAt(2).Ranges, Has.Count.EqualTo(1));
            });
            Assert.That(wb.DefinedNames.ElementAt(2).Ranges.Single().RangeAddress.ToString(XLReferenceStyle.A1, true), Is.EqualTo("'Sheet 1'!A5:D5"));
        }

        [Test]
        public void NamedRangesFromDeletedSheetAreSavedWithoutAddress()
        {
            // Range address referring to the deleted sheet look like #REF!A1:B2.
            // But workbooks with such references in named ranges Excel considers as broken files.
            // It requires #REF!

            using var ms = new MemoryStream();
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
                Assert.That(wb.DefinedNames.Single().RefersTo, Is.EqualTo("#REF!"));
            }
        }

        [Test]
        public void Only_worksheet_scoped_defined_names_are_copied_when_sheet_is_copied()
        {
            using var wb = new XLWorkbook();
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

            Assert.Multiple(() =>
            {
                Assert.That(ws1.Cell("C1").Value, Is.EqualTo(1));
                Assert.That(ws1.Cell("C2").Value, Is.EqualTo(3));
                Assert.That(ws1.Cell("C3").Value, Is.EqualTo(104));
            });

            var wsCopy = ws1.CopyTo("Copy");
            Assert.Multiple(() =>
            {
                Assert.That(wsCopy.Cell("C1").Value, Is.EqualTo(1));
                Assert.That(wsCopy.Cell("C2").Value, Is.EqualTo(3));
                Assert.That(wsCopy.Cell("C3").Value, Is.EqualTo(104));

                Assert.That(wb.DefinedName("wbNamedRange").Ranges.First().RangeAddress.ToStringRelative(true),
                    Is.EqualTo("Sheet1!A1:A10"));
            });
            Assert.That(wsCopy.DefinedName("wsNamedRange").Ranges.First().RangeAddress.ToStringRelative(true),
                Is.EqualTo("Copy!A3:A3"));
            Assert.AreEqual("Sheet2!A4:A4",
                wsCopy.DefinedName("wsNamedRangeAcrossSheets").Ranges.First().RangeAddress.ToStringRelative(true));
        }

        [Test]
        public void Saved_defined_names_become_invalid_on_sheet_deleting()
        {
            using var ms = new MemoryStream();
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
                Assert.Multiple(() =>
                {
                    Assert.That(ws1.DefinedNames.Count(), Is.EqualTo(1));
                    Assert.That(ws1.DefinedNames.First().Name, Is.EqualTo("Named range 1"));
                    Assert.That(ws1.DefinedNames.First().Scope, Is.EqualTo(XLNamedRangeScope.Worksheet));
                    Assert.That(ws1.DefinedNames.First().RefersTo, Is.EqualTo("'Sheet 1'!$A$1:$D$1"));
                    Assert.That(ws1.DefinedNames.First().Ranges.Single().RangeAddress.ToString(XLReferenceStyle.A1, true),
                        Is.EqualTo("'Sheet 1'!A1:D1"));
                });

                Assert.That(wb.DefinedNames.Count(), Is.EqualTo(3));

                Assert.That(wb.DefinedNames.ElementAt(0).Name, Is.EqualTo("Named range 2"));
                Assert.Multiple(() =>
                {
                    Assert.That(wb.DefinedNames.ElementAt(0).Scope, Is.EqualTo(XLNamedRangeScope.Workbook));
                    Assert.That(wb.DefinedNames.ElementAt(0).RefersTo, Is.EqualTo("'Sheet 1'!$A$2:$D$2"));
                    Assert.That(wb.DefinedNames.ElementAt(0).Ranges.Single().RangeAddress.ToString(XLReferenceStyle.A1, true),
                        Is.EqualTo("'Sheet 1'!A2:D2"));
                });

                Assert.That(wb.DefinedNames.ElementAt(1).Name, Is.EqualTo("Named range 4"));
                Assert.That(wb.DefinedNames.ElementAt(1).Scope, Is.EqualTo(XLNamedRangeScope.Workbook));
                Assert.Multiple(() =>
                {
                    Assert.That(wb.DefinedNames.ElementAt(1).RefersTo, Is.EqualTo("#REF!"));
                    Assert.That(wb.DefinedNames.ElementAt(1).Ranges.Any(), Is.False);

                    Assert.That(wb.DefinedNames.ElementAt(2).Name, Is.EqualTo("Named range 5"));
                    Assert.That(wb.DefinedNames.ElementAt(2).Scope, Is.EqualTo(XLNamedRangeScope.Workbook));
                    Assert.That(wb.DefinedNames.ElementAt(2).RefersTo, Is.EqualTo("'Sheet 1'!$A$5:$D$5,#REF!"));
                    Assert.That(wb.DefinedNames.ElementAt(2).Ranges, Has.Count.EqualTo(1));
                });
                Assert.That(wb.DefinedNames.ElementAt(2).Ranges.Single().RangeAddress.ToString(XLReferenceStyle.A1, true),
                    Is.EqualTo("'Sheet 1'!A5:D5"));
            }
        }

        [Test]
        public void TestInvalidNamedRangeOnWorkbookScope()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Column1");
            ws.FirstCell().CellRight().SetValue("Column2").Style.Font.SetBold();
            ws.FirstCell().CellRight(2).SetValue("Column3");

            Assert.Throws<ArgumentException>(() => wb.DefinedNames.Add("MyRange", "A1:C1"));
        }

        [Test]
        public void WbContainsWsNamedRange()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().AddToNamed("Name", XLScope.Worksheet);

            Assert.Multiple(() =>
            {
                Assert.That(wb.DefinedNames.Contains("Sheet1!Name"), Is.True);
                Assert.That(wb.DefinedNames.Contains("Sheet1!NameX"), Is.False);
            });

            Assert.IsNotNull(wb.DefinedName("Sheet1!Name"));
            Assert.That(wb.DefinedName("Sheet1!NameX"), Is.Null);

            Boolean found1 = wb.DefinedNames.TryGetValue("Sheet1!Name", out var definedName1);
            Assert.That(found1, Is.True);
            Assert.IsNotNull(definedName1);
            Assert.That(definedName1.Scope, Is.EqualTo(XLNamedRangeScope.Worksheet));

            Boolean found2 = wb.DefinedNames.TryGetValue("Sheet1!NameX", out var definedName2);
            Assert.That(found2, Is.False);
            Assert.That(definedName2, Is.Null);
        }

        [Test]
        public void WorkbookContainsNamedRange()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().AddToNamed("Name");

            Assert.Multiple(() =>
            {
                Assert.That(wb.DefinedNames.Contains("Name"), Is.True);
                Assert.That(wb.DefinedNames.Contains("NameX"), Is.False);
            });

            Assert.IsNotNull(wb.DefinedName("Name"));
            Assert.That(wb.DefinedName("NameX"), Is.Null);

            Boolean found1 = wb.DefinedNames.TryGetValue("Name", out var definedName1);
            Assert.That(found1, Is.True);
            Assert.IsNotNull(definedName1);

            Boolean found2 = wb.DefinedNames.TryGetValue("NameX", out var definedName2);
            Assert.That(found2, Is.False);
            Assert.That(definedName2, Is.Null);
        }

        [Test]
        public void WorksheetContainsNamedRange()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().AddToNamed("Name", XLScope.Worksheet);

            Assert.Multiple(() =>
            {
                Assert.That(ws.DefinedNames.Contains("Name"), Is.True);
                Assert.That(ws.DefinedNames.Contains("NameX"), Is.False);
            });

            Assert.IsNotNull(ws.DefinedName("Name"));
            Assert.Throws<KeyNotFoundException>(() => ws.DefinedName("NameX"));

            Boolean found1 = ws.DefinedNames.TryGetValue("Name", out var definedName1);
            Assert.That(found1, Is.True);
            Assert.IsNotNull(definedName1);

            Boolean found2 = ws.DefinedNames.TryGetValue("NameX", out var definedName2);
            Assert.That(found2, Is.False);
            Assert.That(definedName2, Is.Null);
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

            Assert.That(a2.GetDouble(), Is.EqualTo(50));
        }
    }
}
