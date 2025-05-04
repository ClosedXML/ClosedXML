using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ClosedXML.Tests
{
    [TestFixture]
    public class XLPivotTableTests
    {
        [Test]
        public void PivotTables()
        {
            Assert.DoesNotThrow(() =>
            {
                using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\PivotTables\PivotTables.xlsx"));
                using var wb = new XLWorkbook(stream);
                var ws = wb.Worksheet("PastrySalesData");
                var table = ws.Table("PastrySalesData");
                var ptSheet = wb.Worksheets.Add("BlankPivotTable");
                ptSheet.PivotTables.Add("pvt", ptSheet.Cell(1, 1), table);

                using var ms = new MemoryStream();
                wb.SaveAs(ms, true);
            });
        }

        [Test]
        public void TestPivotTableVersioningAttributes()
        {
            // Pivot cache definitions in input file has created and refreshed version attributes = 3
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\PivotTableReferenceFiles\VersioningAttributes\inputfile.xlsx"));
            TestHelper.CreateAndCompare(() =>
            {
                var wb = new XLWorkbook(stream);

                var data = wb.Worksheet("Data");

                var pt = data.RangeUsed().CreatePivotTable(wb.AddWorksheet("pvt2").FirstCell(), "pvt2");

                pt.ColumnLabels.Add("Sex");
                pt.RowLabels.Add("FullName");
                pt.Values.Add("Id", "Count of Id").SetSummaryFormula(XLPivotSummary.Count);

                return wb;
                // Pivot cache definitions in output file has created and refreshed version attributes = 5
            }, @"Other\PivotTableReferenceFiles\VersioningAttributes\outputfile.xlsx");
        }

        [Test]
        public void PivotTableOptionsSaveTest()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\PivotTables\PivotTables.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheet("PastrySalesData");
            var table = ws.Table("PastrySalesData");
            var ptSheet = wb.Worksheets.Add("BlankPivotTable");
            var pt = ptSheet.PivotTables.Add("pvtOptionsTest", ptSheet.Cell(1, 1), table);

            pt.ColumnHeaderCaption = "clmn header";
            pt.RowHeaderCaption = "row header";

            pt.AutofitColumns = true;
            pt.PreserveCellFormatting = false;
            pt.ShowGrandTotalsColumns = true;
            pt.ShowGrandTotalsRows = true;
            pt.UseCustomListsForSorting = false;
            pt.ShowExpandCollapseButtons = false;
            pt.ShowContextualTooltips = false;
            pt.DisplayCaptionsAndDropdowns = false;
            pt.RepeatRowLabels = true;
            pt.PivotCache.SaveSourceData = false;
            pt.EnableShowDetails = false;
            pt.ShowColumnHeaders = false;
            pt.ShowRowHeaders = false;

            pt.MergeAndCenterWithLabels = true; // MergeItem
            pt.RowLabelIndent = 12; // Indent
            pt.FilterAreaOrder = XLFilterAreaOrder.OverThenDown; // PageOverThenDown
            pt.FilterFieldsPageWrap = 14; // PageWrap
            pt.ErrorValueReplacement = "error test"; // ErrorCaption
            pt.EmptyCellReplacement = "empty test"; // MissingCaption

            pt.FilteredItemsInSubtotals = true; // Subtotal filtered page items
            pt.AllowMultipleFilters = false; // MultipleFieldFilters

            pt.ShowPropertiesInTooltips = false;
            pt.ClassicPivotTableLayout = true;
            pt.ShowEmptyItemsOnRows = true;
            pt.ShowEmptyItemsOnColumns = true;
            pt.DisplayItemLabels = false;
            pt.SortFieldsAtoZ = true;

            pt.PrintExpandCollapsedButtons = true;
            pt.PrintTitles = true;

            pt.PivotCache.RefreshDataOnOpen = false;
            pt.PivotCache.ItemsToRetainPerField = XLItemsToRetain.Max;
            pt.EnableCellEditing = true;
            pt.ShowValuesRow = true;
            pt.ShowRowStripes = true;
            pt.ShowColumnStripes = true;
            pt.Theme = XLPivotTableTheme.PivotStyleDark13;

            using var ms = new MemoryStream();
            wb.SaveAs(ms, true);

            ms.Position = 0;

            using var wbassert = new XLWorkbook(ms);
            var wsassert = wbassert.Worksheet("BlankPivotTable");
            var ptassert = wsassert.PivotTable("pvtOptionsTest");
            Assert.That(ptassert, Is.Not.Null, "name save failure");
            Assert.Multiple(() =>
            {
                Assert.That(ptassert.ColumnHeaderCaption, Is.EqualTo("clmn header"), "ColumnHeaderCaption save failure");
                Assert.That(ptassert.RowHeaderCaption, Is.EqualTo("row header"), "RowHeaderCaption save failure");
                Assert.That(ptassert.MergeAndCenterWithLabels, Is.True, "MergeAndCenterWithLabels save failure");
                Assert.That(ptassert.RowLabelIndent, Is.EqualTo(12), "RowLabelIndent save failure");
                Assert.That(ptassert.FilterAreaOrder, Is.EqualTo(XLFilterAreaOrder.OverThenDown), "FilterAreaOrder save failure");
                Assert.That(ptassert.FilterFieldsPageWrap, Is.EqualTo(14), "FilterFieldsPageWrap save failure");
                Assert.That(ptassert.ErrorValueReplacement, Is.EqualTo("error test"), "ErrorValueReplacement save failure");
                Assert.That(ptassert.EmptyCellReplacement, Is.EqualTo("empty test"), "EmptyCellReplacement save failure");
                Assert.That(ptassert.AutofitColumns, Is.True, "AutofitColumns save failure");
                Assert.That(ptassert.PreserveCellFormatting, Is.False, "PreserveCellFormatting save failure");
                Assert.That(ptassert.ShowGrandTotalsRows, Is.True, "ShowGrandTotalsRows save failure");
                Assert.That(ptassert.ShowGrandTotalsColumns, Is.True, "ShowGrandTotalsColumns save failure");
                Assert.That(ptassert.FilteredItemsInSubtotals, Is.True, "FilteredItemsInSubtotals save failure");
                Assert.That(ptassert.AllowMultipleFilters, Is.False, "AllowMultipleFilters save failure");
                Assert.That(ptassert.UseCustomListsForSorting, Is.False, "UseCustomListsForSorting save failure");
                Assert.That(ptassert.ShowExpandCollapseButtons, Is.False, "ShowExpandCollapseButtons save failure");
                Assert.That(ptassert.ShowContextualTooltips, Is.False, "ShowContextualTooltips save failure");
                Assert.That(ptassert.ShowPropertiesInTooltips, Is.False, "ShowPropertiesInTooltips save failure");
                Assert.That(ptassert.DisplayCaptionsAndDropdowns, Is.False, "DisplayCaptionsAndDropdowns save failure");
                Assert.That(ptassert.ClassicPivotTableLayout, Is.True, "ClassicPivotTableLayout save failure");
                Assert.That(ptassert.ShowEmptyItemsOnRows, Is.True, "ShowEmptyItemsOnRows save failure");
                Assert.That(ptassert.ShowEmptyItemsOnColumns, Is.True, "ShowEmptyItemsOnColumns save failure");
                Assert.That(ptassert.DisplayItemLabels, Is.False, "DisplayItemLabels save failure");
                Assert.That(ptassert.SortFieldsAtoZ, Is.True, "SortFieldsAtoZ save failure");
                Assert.That(ptassert.PrintExpandCollapsedButtons, Is.True, "PrintExpandCollapsedButtons save failure");
                Assert.That(ptassert.RepeatRowLabels, Is.True, "RepeatRowLabels save failure");
                Assert.That(ptassert.PrintTitles, Is.True, "PrintTitles save failure");
                Assert.That(ptassert.PivotCache.SaveSourceData, Is.False, "SaveSourceData save failure");
                Assert.That(ptassert.EnableShowDetails, Is.False, "EnableShowDetails save failure");
                Assert.That(ptassert.PivotCache.RefreshDataOnOpen, Is.False, "RefreshDataOnOpen save failure");
                Assert.That(ptassert.PivotCache.ItemsToRetainPerField, Is.EqualTo(XLItemsToRetain.Max), "ItemsToRetainPerField save failure");
                Assert.That(ptassert.EnableCellEditing, Is.True, "EnableCellEditing save failure");
                Assert.That(ptassert.Theme, Is.EqualTo(XLPivotTableTheme.PivotStyleDark13), "Theme save failure");
                Assert.That(ptassert.ShowValuesRow, Is.True, "ShowValuesRow save failure");
                Assert.That(ptassert.ShowRowHeaders, Is.False, "ShowRowHeaders save failure");
                Assert.That(ptassert.ShowColumnHeaders, Is.False, "ShowColumnHeaders save failure");
                Assert.That(ptassert.ShowRowStripes, Is.True, "ShowRowStripes save failure");
                Assert.That(ptassert.ShowColumnStripes, Is.True, "ShowColumnStripes save failure");
            });
        }

        [TestCase(true)]
        [TestCase(false)]
        public void PivotFieldOptionsSaveTest(bool withDefaults)
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\PivotTables\PivotTables.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheet("PastrySalesData");
            var table = ws.Table("PastrySalesData");

            var ptSheet = wb.Worksheets.Add("pvtFieldOptionsTest");
            var pt = ptSheet.PivotTables.Add("pvtFieldOptionsTest", ptSheet.Cell(1, 1), table);

            var field = pt.RowLabels.Add("Name")
                .SetSubtotalCaption("Test caption")
                .SetCustomName("Test name");
            SetFieldOptions(field, withDefaults);

            pt.ColumnLabels.Add("Month");
            pt.Values.Add("NumberOfOrders").SetSummaryFormula(XLPivotSummary.Sum);

            using var ms = new MemoryStream();
            wb.SaveAs(ms, true);

            ms.Position = 0;

            using var wbassert = new XLWorkbook(ms);
            var wsassert = wbassert.Worksheet("pvtFieldOptionsTest");
            var ptassert = wsassert.PivotTable("pvtFieldOptionsTest");
            var pfassert = ptassert.RowLabels.Get("Name");
            Assert.That(pfassert, Is.Not.Null, "name save failure");
            Assert.Multiple(() =>
            {
                Assert.That(pfassert.SubtotalCaption, Is.EqualTo("Test caption"), "SubtotalCaption save failure");
                Assert.That(pfassert.CustomName, Is.EqualTo("Test name"), "CustomName save failure");
            });
            AssertFieldOptions(pfassert, withDefaults);
        }

        [Test]
        [Ignore("PT styles will be fixed in a different PR")]
        public void PivotTableStyleFormatsTest()
        {
/*
            using (var ms = new MemoryStream())
            {
                using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\PivotTables\PivotTables.xlsx")))
                using (var wbSource = new XLWorkbook(stream))
                using (var wbDestination = new XLWorkbook())
                {
                    var ws = wbSource.Worksheet("PastrySalesData");
                    wbDestination.AddWorksheet(ws);
                    ws = wbDestination.Worksheet("PastrySalesData");

                    var table = ws.Table("PastrySalesData");
                    var ptSheet = wbDestination.Worksheets.Add("PivotTableStyleFormats");
                    var pt = ptSheet.PivotTables.Add("pvtStyleFormats", ptSheet.Cell(1, 1), table);
                    pt.Layout = XLPivotLayout.Tabular;

                    pt.SetSubtotals(XLPivotSubtotals.AtBottom);

                    var monthPivotField = pt.ColumnLabels.Add("Month");

                    var namePivotField = pt.RowLabels.Add("Name")
                        .SetSubtotalCaption("Test caption")
                        .SetCustomName("Test name")
                        .AddSubtotal(XLSubtotalFunction.Sum);

                    ptSheet.SetTabActive();

                    var numberOfOrdersPivotValue = pt.Values.Add("NumberOfOrders")
                        .SetSummaryFormula(XLPivotSummary.Sum);

                    var qualityPivotValue = pt.Values.Add("Quality").SetSummaryFormula(XLPivotSummary.Sum);

                    pt.StyleFormats.RowGrandTotalFormats.ForElement(XLPivotStyleFormatElement.All).Style.Font.FontColor = XLColor.VenetianRed;

                    namePivotField.StyleFormats.Subtotal.Style.Fill.BackgroundColor = XLColor.Blue;
                    monthPivotField.StyleFormats.Label.Style.Fill.BackgroundColor = XLColor.Amber;
                    monthPivotField.StyleFormats.Header.Style.Font.FontColor = XLColor.Yellow;
                    namePivotField.StyleFormats.DataValuesFormat
                        .AndWith(monthPivotField, v => v.IsText && v.GetText() == "May")
                        .ForValueField(numberOfOrdersPivotValue)
                        .Style.Font.FontColor = XLColor.Green;

                    wbDestination.SaveAs(ms);
                }

                ms.Seek(0, SeekOrigin.Begin);

                using (var wb = new XLWorkbook(ms))
                {
                    var ws = wb.Worksheet("PivotTableStyleFormats");
                    var pt = ws.PivotTable("pvtStyleFormats").CastTo<XLPivotTable>();

                    Assert.AreEqual(0, pt.StyleFormats.ColumnGrandTotalFormats.Count());

                    Assert.NotNull(pt.StyleFormats.RowGrandTotalFormats);
                    Assert.AreEqual(1, pt.StyleFormats.RowGrandTotalFormats.Count());
                    Assert.AreEqual(XLPivotStyleFormatElement.All, pt.StyleFormats.RowGrandTotalFormats.First().AppliesTo);
                    Assert.AreEqual(XLColor.VenetianRed, pt.StyleFormats.RowGrandTotalFormats.ForElement(XLPivotStyleFormatElement.All).Style.Font.FontColor);

                    var namePivotField = pt.RowLabels.Get("Name");
                    var monthPivotField = pt.ColumnLabels.Get("Month");
                    var numberOfOrdersPivotValue = pt.Values.Get("NumberOfOrders");

                    Assert.AreEqual(XLStyle.Default, namePivotField.StyleFormats.Label.Style);
                    Assert.AreEqual(XLColor.Blue, namePivotField.StyleFormats.Subtotal.Style.Fill.BackgroundColor);

                    Assert.AreEqual(XLStyle.Default, monthPivotField.StyleFormats.Subtotal.Style);
                    Assert.AreEqual(XLColor.Amber, monthPivotField.StyleFormats.Label.Style.Fill.BackgroundColor);
                    Assert.AreEqual(XLColor.Yellow, monthPivotField.StyleFormats.Header.Style.Font.FontColor);

                    var nameDataValuesFormat = namePivotField.StyleFormats.DataValuesFormat as XLPivotValueStyleFormat;
                    Assert.AreEqual(2, nameDataValuesFormat.FieldReferences.Count());

                    Assert.AreEqual(monthPivotField, nameDataValuesFormat.FieldReferences.First().CastTo<PivotLabelFieldReference>().PivotField);

                    Assert.AreEqual(numberOfOrdersPivotValue.CustomName, nameDataValuesFormat.FieldReferences.Last().CastTo<PivotValueFieldReference>().Value);

                    wb.Save();
                }
            }
*/
        }

        [Test]
        public void CopyPivotTableTests()
        {
            using var ms = new MemoryStream();
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\PivotTables\PivotTables.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws1 = wb.Worksheet("pvt1");
            var pt1 = ws1.PivotTables.First() as XLPivotTable;

            Assert.Throws<InvalidOperationException>(() => pt1.CopyTo(pt1.TargetCell));

            var pt2 = pt1.CopyTo(ws1.Cell("AB100")) as XLPivotTable;

            AssertPivotTablesAreEqual(pt1, pt2, compareName: false);

            var ws2 = wb.AddWorksheet("Copy Of pvt1");
            AssertPivotTablesAreEqual(pt1, pt1.CopyTo(ws2.FirstCell()) as XLPivotTable, compareName: true);

            using var wb2 = new XLWorkbook();
            wb.Worksheet("PastrySalesData").CopyTo(wb2);

            AssertPivotTablesAreEqual(pt1, pt1.CopyTo(wb2.AddWorksheet("pvt").FirstCell()) as XLPivotTable, compareName: true);
        }

        private void AssertPivotTablesAreEqual(XLPivotTable original, XLPivotTable copy, bool compareName)
        {
            Assert.That(original.Name.Equals(copy.Name), Is.EqualTo(compareName));

            var comparer = new PivotTableComparer(compareName: compareName, compareRelId: false, compareTargetCellAddress: false);
            Assert.That(comparer.Equals(original, copy), Is.True);
        }

        private class Pastry
        {
            public Pastry(string name, int? code, int numberOfOrders, double quality, string month, DateTime? bakeDate)
            {
                Name = name;
                Code = code;
                NumberOfOrders = numberOfOrders;
                Quality = quality;
                Month = month;
                BakeDate = bakeDate;
            }

            public string Name { get; set; }
            public int? Code { get; }
            public int NumberOfOrders { get; set; }
            public double Quality { get; set; }
            public string Month { get; set; }
            public DateTime? BakeDate { get; set; }
        }

        [Test]
        public void SharedItemsWithVariousDataTypesInTableColumn()
        {
            // Load an excel that contains a table which has various combinations of types in columns.
            // The pivot cache definition contain various flags in shared items for each field and the
            // test checks the flags in cache are set correctly (they are determined in cache writer).
            TestHelper.LoadSaveAndCompare(
                @"Other\PivotTableReferenceFiles\VariousDataTypesInTableColumns\input.xlsx",
                @"Other\PivotTableReferenceFiles\VariousDataTypesInTableColumns\output.xlsx");
        }

        [Test]
        public void BlankPivotTableField()
        {
            using var ms = new MemoryStream();
            TestHelper.CreateAndCompare(() =>
            {
                // Based on .\ClosedXML\ClosedXML.Examples\PivotTables\PivotTables.cs
                // But with empty column for Month
                var pastries = new List<Pastry>
                {
                    new Pastry("Croissant", 101, 150, 60.2, "", new DateTime(2016, 04, 21)),
                    new Pastry("Croissant", 101, 250, 50.42, "", new DateTime(2016, 05, 03)),
                    new Pastry("Croissant", 101, 134, 22.12, "", new DateTime(2016, 06, 24)),
                    new Pastry("Doughnut", 102, 250, 89.99, "", new DateTime(2017, 04, 23)),
                    new Pastry("Doughnut", 102, 225, 70, "", new DateTime(2016, 05, 24)),
                    new Pastry("Doughnut", 102, 210, 75.33, "", new DateTime(2016, 06, 02)),
                    new Pastry("Bearclaw", 103, 134, 10.24, "", new DateTime(2016, 04, 27)),
                    new Pastry("Bearclaw", 103, 184, 33.33, "", new DateTime(2016, 05, 20)),
                    new Pastry("Bearclaw", 103, 124, 25, "", new DateTime(2017, 06, 05)),
                    new Pastry("Danish", 104, 394, -20.24, "", null),
                    new Pastry("Danish", 104, 190, 60, "", new DateTime(2017, 05, 08)),
                    new Pastry("Danish", 104, 221, 24.76, "", new DateTime(2016, 06, 21)),

                    // Deliberately add different casings of same string to ensure pivot table doesn't duplicate it.
                    new Pastry("Scone", 105, 135, 0, "", new DateTime(2017, 04, 22)),
                    new Pastry("SconE", 105, 122, 5.19, "", new DateTime(2017, 05, 03)),
                    new Pastry("SCONE", 105, 243, 44.2, "", new DateTime(2017, 06, 14)),

                    // For ContainsBlank and integer rows/columns test
                    new Pastry("Scone", null, 255, 18.4, "", null),
                };

                var wb = new XLWorkbook();

                var sheet = wb.Worksheets.Add("PastrySalesData");
                // Insert our list of pastry data into the "PastrySalesData" sheet at cell 1,1
                var table = sheet.Cell(1, 1).InsertTable(pastries, "PastrySalesData", true);
                sheet.Cell("F11").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                sheet.Columns().AdjustToContents();

                IXLWorksheet ptSheet;
                IXLPivotTable pt;

                for (var i = 1; i <= 5; i++)
                {
                    // Add a new sheet for our pivot table
                    ptSheet = wb.Worksheets.Add("pvt" + i);

                    // Create the pivot table, using the data from the "PastrySalesData" table
                    pt = ptSheet.PivotTables.Add("pvt" + i, ptSheet.Cell(1, 1), table);

                    if (i == 1 || i == 4 || i == 5)
                        pt.ColumnLabels.Add("Name");
                    else if (i == 2 || i == 3)
                        pt.RowLabels.Add("Name");

                    if (i == 1 || i == 3)
                        pt.RowLabels.Add("Month");
                    else if (i == 2 || i == 4)
                        pt.ColumnLabels.Add("Month");
                    else if (i == 5)
                        pt.RowLabels.Add("BakeDate");

                    // The values in our table will come from the "NumberOfOrders" field
                    // The default calculation setting is a total of each row/column
                    pt.Values.Add("NumberOfOrders", "NumberOfOrdersPercentageOfBearclaw")
                        .ShowAsPercentageFrom("Name").And("Bearclaw")
                        .NumberFormat.Format = "0%";

                    ptSheet.Columns().AdjustToContents();
                }

                return wb;
            }, @"Other\PivotTableReferenceFiles\BlankPivotTableField\BlankPivotTableField.xlsx");
        }

        [Test]
        public void SourceSheetWithWhitespace()
        {
            // Check that pivot source reference for a sheet name with whitespaces
            // is not saved to the file with escaped quotes, issue #955.
            TestHelper.CreateAndCompare(() =>
            {
                var wb = new XLWorkbook();

                // Worksheet name contains whitespaces that shouldn't be quoted in the file.
                var sheet = wb.Worksheets.Add("Pastry Sales Data");
                var range = sheet.Cell(1, 1).InsertData(new object[]
                {
                    ("Name", "Sold count"),
                    ("Pie", 7),
                    ("Cake", 10),
                    ("Pie", 2),
                });

                // Add a new sheet for our pivot table
                var ptSheet = wb.Worksheets.Add("pvt");
                var pt = ptSheet.PivotTables.Add("pvt", ptSheet.Cell(1, 1), range);
                pt.RowLabels.Add("Name");
                pt.Values.Add("Sold count");

                return wb;
            }, @"Other\PivotTableReferenceFiles\SourceSheetWithWhitespace\outputfile.xlsx");
        }

        [Test]
        public void PivotTableWithNoneTheme()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\PivotTableReferenceFiles\PivotTableWithNoneTheme\inputfile.xlsx"));
            using var ms = new MemoryStream();
            TestHelper.CreateAndCompare(() =>
            {
                var wb = new XLWorkbook(stream);
                wb.SaveAs(ms);
                return wb;
            }, @"Other\PivotTableReferenceFiles\PivotTableWithNoneTheme\outputfile.xlsx");
        }

        [Test]
        public void MaintainPivotTableLabelsOrder()
        {
            var pastries = new List<Pastry>
            {
                new Pastry("Croissant", 101, 150, 60.2, "", new DateTime(2016, 04, 21)),
                new Pastry("Croissant", 101, 250, 50.42, "", new DateTime(2016, 05, 03)),
                new Pastry("Croissant", 101, 134, 22.12, "", new DateTime(2016, 06, 24)),
                new Pastry("Doughnut", 102, 250, 89.99, "", new DateTime(2017, 04, 23)),
                new Pastry("Doughnut", 102, 225, 70, "", new DateTime(2016, 05, 24)),
                new Pastry("Doughnut", 102, 210, 75.33, "", new DateTime(2016, 06, 02)),
                new Pastry("Bearclaw", 103, 134, 10.24, "", new DateTime(2016, 04, 27)),
                new Pastry("Bearclaw", 103, 184, 33.33, "", new DateTime(2016, 05, 20)),
                new Pastry("Bearclaw", 103, 124, 25, "", new DateTime(2017, 06, 05)),
                new Pastry("Danish", 104, 394, -20.24, "", null),
                new Pastry("Danish", 104, 190, 60, "", new DateTime(2017, 05, 08)),
                new Pastry("Danish", 104, 221, 24.76, "", new DateTime(2016, 06, 21)),

                // Deliberately add different casings of same string to ensure pivot table doesn't duplicate it.
                new Pastry("Scone", 105, 135, 0, "", new DateTime(2017, 04, 22)),
                new Pastry("SconE", 105, 122, 5.19, "", new DateTime(2017, 05, 03)),
                new Pastry("SCONE", 105, 243, 44.2, "", new DateTime(2017, 06, 14)),

                // For ContainsBlank and integer rows/columns test
                new Pastry("Scone", null, 255, 18.4, "", null),
            };

            using (var ms = new MemoryStream())
            {
                // Page fields
                using (var wb = new XLWorkbook())
                {
                    var sheet = wb.Worksheets.Add("PastrySalesData");
                    // Insert our list of pastry data into the "PastrySalesData" sheet at cell 1,1
                    var table = sheet.Cell(1, 1).InsertTable(pastries, "PastrySalesData", true);
                    sheet.Cell("F11").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    sheet.Columns().AdjustToContents();

                    IXLWorksheet ptSheet;
                    IXLPivotTable pt;

                    // Add a new sheet for our pivot table
                    ptSheet = wb.Worksheets.Add("pvt");

                    // Create the pivot table, using the data from the "PastrySalesData" table
                    pt = ptSheet.PivotTables.Add("PastryPivot", ptSheet.Cell(1, 1), table);

                    pt.ReportFilters.Add("Month");
                    pt.ReportFilters.Add("Name");

                    pt.RowLabels.Add("BakeDate");
                    pt.Values.Add("NumberOfOrders").SetSummaryFormula(XLPivotSummary.Sum);

                    wb.SaveAs(ms);
                }

                ms.Seek(0, SeekOrigin.Begin);

                using (var wb = new XLWorkbook(ms))
                {
                    var pageFields = wb.Worksheets.SelectMany(ws => ws.PivotTables)
                        .First()
                        .ReportFilters
                        .ToArray();

                    Assert.Multiple(() =>
                    {
                        Assert.That(pageFields[0].SourceName, Is.EqualTo("Month"));
                        Assert.That(pageFields[1].SourceName, Is.EqualTo("Name"));
                    });
                }
            }

            using (var ms = new MemoryStream())
            {
                // Column labels
                using (var wb = new XLWorkbook())
                {
                    var sheet = wb.Worksheets.Add("PastrySalesData");
                    // Insert our list of pastry data into the "PastrySalesData" sheet at cell 1,1
                    var table = sheet.Cell(1, 1).InsertTable(pastries, "PastrySalesData", true);
                    sheet.Cell("F11").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    sheet.Columns().AdjustToContents();

                    IXLWorksheet ptSheet;
                    IXLPivotTable pt;

                    // Add a new sheet for our pivot table
                    ptSheet = wb.Worksheets.Add("pvt");

                    // Create the pivot table, using the data from the "PastrySalesData" table
                    pt = ptSheet.PivotTables.Add("PastryPivot", ptSheet.Cell(1, 1), table);

                    pt.ColumnLabels.Add("Month");
                    pt.ColumnLabels.Add("Name");

                    pt.RowLabels.Add("BakeDate");
                    pt.Values.Add("NumberOfOrders").SetSummaryFormula(XLPivotSummary.Sum);

                    wb.SaveAs(ms);
                }

                ms.Seek(0, SeekOrigin.Begin);

                using (var wb = new XLWorkbook(ms))
                {
                    var columnLabels = wb.Worksheets.SelectMany(ws => ws.PivotTables)
                        .First()
                        .ColumnLabels
                        .ToArray();

                    Assert.Multiple(() =>
                    {
                        Assert.That(columnLabels[0].SourceName, Is.EqualTo("Month"));
                        Assert.That(columnLabels[1].SourceName, Is.EqualTo("Name"));
                    });
                }
            }

            using (var ms = new MemoryStream())
            {
                // Row labels
                using (var wb = new XLWorkbook())
                {
                    var sheet = wb.Worksheets.Add("PastrySalesData");
                    // Insert our list of pastry data into the "PastrySalesData" sheet at cell 1,1
                    var table = sheet.Cell(1, 1).InsertTable(pastries, "PastrySalesData", true);
                    sheet.Cell("F11").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    sheet.Columns().AdjustToContents();

                    IXLWorksheet ptSheet;
                    IXLPivotTable pt;

                    // Add a new sheet for our pivot table
                    ptSheet = wb.Worksheets.Add("pvt");

                    // Create the pivot table, using the data from the "PastrySalesData" table
                    pt = ptSheet.PivotTables.Add("PastryPivot", ptSheet.Cell(1, 1), table);

                    pt.RowLabels.Add("Month");
                    pt.RowLabels.Add("Name");
                    pt.RowLabels.Add(XLConstants.PivotTable.ValuesSentinalLabel);

                    pt.ColumnLabels.Add("BakeDate");
                    pt.Values.Add("NumberOfOrders").SetSummaryFormula(XLPivotSummary.Sum);

                    wb.SaveAs(ms);
                }

                ms.Seek(0, SeekOrigin.Begin);

                using (var wb = new XLWorkbook(ms))
                {
                    var rowLabels = wb.Worksheets.SelectMany(ws => ws.PivotTables)
                        .First()
                        .RowLabels
                        .ToArray();

                    Assert.Multiple(() =>
                    {
                        Assert.That(rowLabels[0].SourceName, Is.EqualTo("Month"));
                        Assert.That(rowLabels[1].SourceName, Is.EqualTo("Name"));
                        Assert.That(rowLabels[2].SourceName, Is.EqualTo("{{Values}}"));
                    });
                }
            }
        }

        [Test]
        public void MaintainPivotTableIntegrityOnMultipleSaves()
        {
            var pastries = new List<Pastry>
            {
                new Pastry("Croissant", 101, 150, 60.2, "", new DateTime(2016, 04, 21)),
                new Pastry("Croissant", 101, 250, 50.42, "", new DateTime(2016, 05, 03)),
                new Pastry("Croissant", 101, 134, 22.12, "", new DateTime(2016, 06, 24)),
                new Pastry("Doughnut", 102, 250, 89.99, "", new DateTime(2017, 04, 23)),
                new Pastry("Doughnut", 102, 225, 70, "", new DateTime(2016, 05, 24)),
                new Pastry("Doughnut", 102, 210, 75.33, "", new DateTime(2016, 06, 02)),
                new Pastry("Bearclaw", 103, 134, 10.24, "", new DateTime(2016, 04, 27)),
                new Pastry("Bearclaw", 103, 184, 33.33, "", new DateTime(2016, 05, 20)),
                new Pastry("Bearclaw", 103, 124, 25, "", new DateTime(2017, 06, 05)),
                new Pastry("Danish", 104, 394, -20.24, "", null),
                new Pastry("Danish", 104, 190, 60, "", new DateTime(2017, 05, 08)),
                new Pastry("Danish", 104, 221, 24.76, "", new DateTime(2016, 06, 21)),

                // Deliberately add different casings of same string to ensure pivot table doesn't duplicate it.
                new Pastry("Scone", 105, 135, 0, "", new DateTime(2017, 04, 22)),
                new Pastry("SconE", 105, 122, 5.19, "", new DateTime(2017, 05, 03)),
                new Pastry("SCONE", 105, 243, 44.2, "", new DateTime(2017, 06, 14)),

                // For ContainsBlank and integer rows/columns test
                new Pastry("Scone", null, 255, 18.4, "", null),
            };

            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("PastrySalesData");
                var table = ws.FirstCell().InsertTable(pastries, "PastrySalesData", true);

                var pvtSheet = wb.Worksheets.Add("pvt");
                var pvt = table.CreatePivotTable(pvtSheet.FirstCell(), "PastryPvt");

                pvt.ColumnLabels.Add("Month");
                pvt.RowLabels.Add("Name");
                pvt.Values.Add("NumberOfOrders").SetSummaryFormula(XLPivotSummary.Sum);

                //Deliberately try to save twice
                wb.SaveAs(ms);
                wb.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (var wb = new XLWorkbook(ms))
            {
                Assert.That(wb.Worksheets.SelectMany(ws => ws.PivotTables).Count(), Is.EqualTo(1));
            }
        }

        [Test]
        public void TwoPivotWithOneSourceTest()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\PivotTableReferenceFiles\TwoPivotTablesWithSingleSource\input.xlsx"));
            TestHelper.CreateAndCompare(() =>
            {
                var wb = new XLWorkbook(stream);
                var srcRange = wb.Range("Sheet1!$B$2:$H$207");

                var pivotSource = wb.PivotCaches.Add(srcRange);

                foreach (var pt in wb.Worksheets.SelectMany(ws => ws.PivotTables))
                {
                    pt.PivotCache = pivotSource;
                }

                return wb;
            }, @"Other\PivotTableReferenceFiles\TwoPivotTablesWithSingleSource\output.xlsx");
        }

        [Test]
        public void PivotSubtotalsLoadingTest()
        {
            // Make sure that if the original file has *subtotals*, the subtotals are
            // turned on even after loading into ClosedXML and then saving the document.
            TestHelper.LoadSaveAndCompare(
                @"Other\PivotTableReferenceFiles\PivotSubtotalsSource\input.xlsx",
                @"Other\PivotTableReferenceFiles\PivotSubtotalsSource\output.xlsx");
        }

        [Test]
        public void ClearPivotTableRenderedRange()
        {
            // https://github.com/ClosedXML/ClosedXML/pull/856
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\PivotTableReferenceFiles\ClearPivotTableRenderedRangeWhenLoading\inputfile.xlsx"));
            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheet("Sheet1");
                Assert.Multiple(() =>
                {
                    Assert.That(ws.Cell("B1").IsEmpty(), Is.True);
                    Assert.That(ws.Cell("C2").IsEmpty(), Is.True);
                    Assert.That(ws.Cell("D5").IsEmpty(), Is.True);
                });
                wb.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (var wb = new XLWorkbook(ms))
            {
                var ws = wb.Worksheet("Sheet1");
                Assert.Multiple(() =>
                {
                    Assert.That(ws.Cell("B1").IsEmpty(), Is.True);
                    Assert.That(ws.Cell("C2").IsEmpty(), Is.True);
                    Assert.That(ws.Cell("D5").IsEmpty(), Is.True);
                });
            }
        }

        [Test]
        public void Add_all_pivot_tables_for_same_range_use_same_pivot_cache()
        {
            // Two different pivot tables created from same range use same pivot cache
            // and don't create a separate pivot cache for each pivot table.
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var range = ws.FirstCell().InsertData(new object[]
            {
                ("Name", "Count"),
                ("Pie", 14),
            });

            var rangePivot1 = ws.PivotTables.Add("rangePivot1", ws.Cell("D1"), range);
            var rangePivot2 = ws.PivotTables.Add("rangePivot2", ws.Cell("D20"), range);

            Assert.That(rangePivot2, Is.Not.SameAs(rangePivot1));
            Assert.That(rangePivot2.PivotCache, Is.SameAs(rangePivot1.PivotCache));
        }

        [Test]
        public void Add_all_pivot_tables_for_same_table_use_same_pivot_cache()
        {
            // Two different pivot tables created from same table use same pivot cache
            // and don't create a separate pivot cache for each pivot table.
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var table = ws.FirstCell().InsertTable(new object[]
            {
                ("Name", "Count"),
                ("Pie", 14),
            });

            var tablePivot1 = ws.PivotTables.Add("tablePivot1", ws.Cell("J1"), table);
            var tablePivot2 = ws.PivotTables.Add("tablePivot2", ws.Cell("J20"), table);

            Assert.That(tablePivot2, Is.Not.SameAs(tablePivot1));
            Assert.That(tablePivot2.PivotCache, Is.SameAs(tablePivot1.PivotCache));
        }

        [Test]
        public void Add_pivot_tables_will_use_table_as_source_if_range_matches_table_area()
        {
            // When a pivot table is created, the `Add` method tries to first
            // find a table with same area as the requested range. If it finds one,
            // the cache will be created from the table and not a range. That is the
            // Excel behavior and generally makes sense.
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.FirstCell().InsertTable(new object[]
            {
                ("Name", "Count"),
                ("Pie", 14),
            }, "Test table");

            // A range that matches the size of an area
            var matchingRange = ws.Range("A1:B3");

            var tablePivot1 = ws.PivotTables.Add("tablePivot1", ws.Cell("J1"), matchingRange);

            var cacheSource = (XLPivotSourceReference)((XLPivotCache)tablePivot1.PivotCache).Source;
            Assert.Multiple(() =>
            {
                Assert.That(cacheSource.UsesName, Is.True);
                Assert.That(cacheSource.Name, Is.EqualTo("Test table"));
            });
        }

        [Test]
        public void Load_and_save_pivot_table_with_cache_records_but_missing_source_data()
        {
            // Test file contains a pivot table created from a normal table in
            // a sheet that was already deleted. The file contains cache records,
            // but the table and original sheet are gone. It's possible to load
            // and save such a pivot table.
            // Opening the saved file in Excel throws an error 'Reference isn't valid'
            // on load, because of `RefreshOnLoad` flag. That flag is always enabled because
            // ClosedXML relies on Excel to rebuild the table and fix it.
            // At this time, there is no content, only shape, because we don't have an engine
            // to determine correct layout and values. Change RefreshDataOnOpen to 0 and change
            // PT in Excel to see the values (aka gimp on Excel PT engine).
            TestHelper.LoadSaveAndCompare(
                @"Other\PivotTableReferenceFiles\PivotTableWithoutSourceData-input.xlsx",
                @"Other\PivotTableReferenceFiles\PivotTableWithoutSourceData-output.xlsx");
        }

        [Test]
        public void Skips_chartsheets_during_pivot_table_loading()
        {
            // Pivot table loading code looks for pivot tables on each sheet, but it shouldn't
            // crash when sheet is a chartsheet or other type of sheet. The referenced test file
            // contains chartsheet and a pivot table to ensure that loading code won't crash.
            TestHelper.LoadAndAssert(wb =>
            {
                // Check that existing pivot table is loaded.
                Assert.That(wb.Worksheet("pivot").PivotTables.Contains("Pastries"), Is.True);
            }, @"Other\PivotTableReferenceFiles\ChartsheetAndPivotTable.xlsx");
        }

        #region IXLPivotTable properties

        #region TargetCell

        [Test]
        public void Property_TargetCell_sets_value_of_the_top_left_corner_of_pivot_table()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var data = ws.Cell("A1").InsertData(new object[]
            {
                ("Name", "City", "Flavor", "Sales"),
                ("Cake", "Tokyo", "Vanilla", 7),
            });
            var pt = ws.PivotTables.Add("pt", ws.Cell("E1"), data);
            pt.ReportFilters.Add("City");

            Assert.Multiple(() =>
            {
                // Even when we added filter and a gap row, the target cell is still E1
                Assert.That(pt.TargetCell.Address.ToString(), Is.EqualTo("E1"));
                Assert.That(((XLPivotTable)pt).Area.FirstPoint.ToString(), Is.EqualTo("E3"));
            });

            pt.TargetCell = ws.Cell("E2");
            Assert.Multiple(() =>
            {
                Assert.That(pt.TargetCell.Address.ToString(), Is.EqualTo("E2"));
                Assert.That(((XLPivotTable)pt).Area.FirstPoint.ToString(), Is.EqualTo("E4"));
            });
        }

        #endregion

        #region FilterAreaOrder

        [TestCase(XLFilterAreaOrder.DownThenOver, "E5")]
        [TestCase(XLFilterAreaOrder.OverThenDown, "E3")]
        public void Property_FilterAreaOrder_determines_direction_in_which_are_filter_fields_laid_out(XLFilterAreaOrder order, string tableAddress)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var data = ws.Cell("A1").InsertData(new object[]
            {
                ("Name", "City", "Flavor", "Sales"),
                ("Cake", "Tokyo", "Vanilla", 7),
            });

            var pt = ws.PivotTables.Add("pt", ws.Cell("E1"), data);
            pt.FilterAreaOrder = order;

            pt.ReportFilters.Add("Name");
            pt.ReportFilters.Add("City");
            pt.ReportFilters.Add("Flavor");

            // Indirect detection of filter fields layout: The address of pivot table are is
            // determined by filter area order.
            Assert.That(((XLPivotTable)pt).Area.ToString(), Is.EqualTo(tableAddress));
        }

        #endregion

        #region Layout

        [TestCase(XLPivotLayout.Outline, "Property_layout_sets_layout_of_pivot_table_and_all_fields-outline.xlsx")]
        [TestCase(XLPivotLayout.Tabular, "Property_layout_sets_layout_of_pivot_table_and_all_fields-tabular.xlsx")]
        [TestCase(XLPivotLayout.Compact, "Property_layout_sets_layout_of_pivot_table_and_all_fields-compact.xlsx")]
        public void Property_layout_sets_layout_of_pivot_table_and_all_fields(XLPivotLayout layout, string testFile)
        {
            // The pivot table also contains unused field Currency. It is there, because tabular
            // layout doesn't display header fields properly (i.e. one header per axis field),
            // unless all (even fields that are not on any axis) have the same field layout.
            TestHelper.CreateAndCompare(wb =>
            {
                var dataSheet = wb.AddWorksheet();
                var dataRange = dataSheet.Cell("A1").InsertData(new object[]
                {
                    ("Name", "Size", "Month", "Season", "Price", "Currency"),
                    ("Cake", "Small", "Jan", "Winter", 9, "EUR"),
                    ("Pie", "Small", "Jan", "Winter", 7, "EUR"),
                    ("Cake", "Large", "Feb", "Summer", 3, "CZK"),
                });

                var ptSheet = wb.AddWorksheet().SetTabActive();
                ptSheet.Column("A").Width = 15;
                var pt = dataRange.CreatePivotTable(ptSheet.Cell("A1"), "pivot table");

                // Add at least two fields to each axis to make each layout distinctive.
                pt.RowLabels.Add("Name");
                pt.RowLabels.Add("Size");
                pt.ColumnLabels.Add("Month");
                pt.ColumnLabels.Add("Season");
                pt.Values.Add("Price");

                pt.Layout = layout;
            }, $@"Other\PivotTable\TableProps\{testFile}");
        }

        #endregion

        #endregion

        private static void SetFieldOptions(IXLPivotField field, bool withDefaults)
        {
            field.SubtotalsAtTop = !withDefaults;
            field.ShowBlankItems = !withDefaults;
            field.Outline = !withDefaults;
            field.Compact = !withDefaults;
            field.Collapsed = withDefaults;
            field.InsertBlankLines = withDefaults;
            field.RepeatItemLabels = withDefaults;
            field.InsertPageBreaks = withDefaults;
            field.IncludeNewItemsInFilter = withDefaults;
        }

        private static void AssertFieldOptions(IXLPivotField field, bool withDefaults)
        {
            Assert.Multiple(() =>
            {
                Assert.That(field.SubtotalsAtTop, Is.EqualTo(!withDefaults), "SubtotalsAtTop save failure");
                Assert.That(field.ShowBlankItems, Is.EqualTo(!withDefaults), "ShowBlankItems save failure");
                Assert.That(field.Outline, Is.EqualTo(!withDefaults), "Outline save failure");
                Assert.That(field.Compact, Is.EqualTo(!withDefaults), "Compact save failure");
                Assert.That(field.Collapsed, Is.EqualTo(withDefaults), "Collapsed save failure");
                Assert.That(field.InsertBlankLines, Is.EqualTo(withDefaults), "InsertBlankLines save failure");
                Assert.That(field.RepeatItemLabels, Is.EqualTo(withDefaults), "RepeatItemLabels save failure");
                Assert.That(field.InsertPageBreaks, Is.EqualTo(withDefaults), "InsertPageBreaks save failure");
                Assert.That(field.IncludeNewItemsInFilter, Is.EqualTo(withDefaults), "IncludeNewItemsInFilter save failure");
            });
        }
    }
}
