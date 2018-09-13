using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ClosedXML_Tests
{
    [TestFixture]
    public class XLPivotTableTests
    {
        [Test]
        public void PivotTables()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\PivotTables\PivotTables.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheet("PastrySalesData");
                var table = ws.Table("PastrySalesData");
                var ptSheet = wb.Worksheets.Add("BlankPivotTable");
                ptSheet.PivotTables.Add("pvt", ptSheet.Cell(1, 1), table);

                using (var ms = new MemoryStream())
                {
                    wb.SaveAs(ms, true);
                }
            }
        }

        [Test]
        public void PivotTableOptionsSaveTest()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\PivotTables\PivotTables.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
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
                pt.SaveSourceData = false;
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

                // TODO pt.RefreshDataOnOpen = false;
                pt.ItemsToRetainPerField = XLItemsToRetain.Max;
                pt.EnableCellEditing = true;
                pt.ShowValuesRow = true;
                pt.ShowRowStripes = true;
                pt.ShowColumnStripes = true;
                pt.Theme = XLPivotTableTheme.PivotStyleDark13;

                using (var ms = new MemoryStream())
                {
                    wb.SaveAs(ms, true);

                    ms.Position = 0;

                    using (var wbassert = new XLWorkbook(ms))
                    {
                        var wsassert = wbassert.Worksheet("BlankPivotTable");
                        var ptassert = wsassert.PivotTable("pvtOptionsTest");
                        Assert.AreNotEqual(null, ptassert, "name save failure");
                        Assert.AreEqual("clmn header", ptassert.ColumnHeaderCaption, "ColumnHeaderCaption save failure");
                        Assert.AreEqual("row header", ptassert.RowHeaderCaption, "RowHeaderCaption save failure");
                        Assert.AreEqual(true, ptassert.MergeAndCenterWithLabels, "MergeAndCenterWithLabels save failure");
                        Assert.AreEqual(12, ptassert.RowLabelIndent, "RowLabelIndent save failure");
                        Assert.AreEqual(XLFilterAreaOrder.OverThenDown, ptassert.FilterAreaOrder, "FilterAreaOrder save failure");
                        Assert.AreEqual(14, ptassert.FilterFieldsPageWrap, "FilterFieldsPageWrap save failure");
                        Assert.AreEqual("error test", ptassert.ErrorValueReplacement, "ErrorValueReplacement save failure");
                        Assert.AreEqual("empty test", ptassert.EmptyCellReplacement, "EmptyCellReplacement save failure");
                        Assert.AreEqual(true, ptassert.AutofitColumns, "AutofitColumns save failure");
                        Assert.AreEqual(false, ptassert.PreserveCellFormatting, "PreserveCellFormatting save failure");
                        Assert.AreEqual(true, ptassert.ShowGrandTotalsRows, "ShowGrandTotalsRows save failure");
                        Assert.AreEqual(true, ptassert.ShowGrandTotalsColumns, "ShowGrandTotalsColumns save failure");
                        Assert.AreEqual(true, ptassert.FilteredItemsInSubtotals, "FilteredItemsInSubtotals save failure");
                        Assert.AreEqual(false, ptassert.AllowMultipleFilters, "AllowMultipleFilters save failure");
                        Assert.AreEqual(false, ptassert.UseCustomListsForSorting, "UseCustomListsForSorting save failure");
                        Assert.AreEqual(false, ptassert.ShowExpandCollapseButtons, "ShowExpandCollapseButtons save failure");
                        Assert.AreEqual(false, ptassert.ShowContextualTooltips, "ShowContextualTooltips save failure");
                        Assert.AreEqual(false, ptassert.ShowPropertiesInTooltips, "ShowPropertiesInTooltips save failure");
                        Assert.AreEqual(false, ptassert.DisplayCaptionsAndDropdowns, "DisplayCaptionsAndDropdowns save failure");
                        Assert.AreEqual(true, ptassert.ClassicPivotTableLayout, "ClassicPivotTableLayout save failure");
                        Assert.AreEqual(true, ptassert.ShowEmptyItemsOnRows, "ShowEmptyItemsOnRows save failure");
                        Assert.AreEqual(true, ptassert.ShowEmptyItemsOnColumns, "ShowEmptyItemsOnColumns save failure");
                        Assert.AreEqual(false, ptassert.DisplayItemLabels, "DisplayItemLabels save failure");
                        Assert.AreEqual(true, ptassert.SortFieldsAtoZ, "SortFieldsAtoZ save failure");
                        Assert.AreEqual(true, ptassert.PrintExpandCollapsedButtons, "PrintExpandCollapsedButtons save failure");
                        Assert.AreEqual(true, ptassert.RepeatRowLabels, "RepeatRowLabels save failure");
                        Assert.AreEqual(true, ptassert.PrintTitles, "PrintTitles save failure");
                        Assert.AreEqual(false, ptassert.SaveSourceData, "SaveSourceData save failure");
                        Assert.AreEqual(false, ptassert.EnableShowDetails, "EnableShowDetails save failure");
                        // TODO Assert.AreEqual(false, ptassert.RefreshDataOnOpen, "RefreshDataOnOpen save failure");
                        Assert.AreEqual(XLItemsToRetain.Max, ptassert.ItemsToRetainPerField, "ItemsToRetainPerField save failure");
                        Assert.AreEqual(true, ptassert.EnableCellEditing, "EnableCellEditing save failure");
                        Assert.AreEqual(XLPivotTableTheme.PivotStyleDark13, ptassert.Theme, "Theme save failure");
                        Assert.AreEqual(true, ptassert.ShowValuesRow, "ShowValuesRow save failure");
                        Assert.AreEqual(false, ptassert.ShowRowHeaders, "ShowRowHeaders save failure");
                        Assert.AreEqual(false, ptassert.ShowColumnHeaders, "ShowColumnHeaders save failure");
                        Assert.AreEqual(true, ptassert.ShowRowStripes, "ShowRowStripes save failure");
                        Assert.AreEqual(true, ptassert.ShowColumnStripes, "ShowColumnStripes save failure");
                    }
                }
            }
        }

        [TestCase(true)]
        [TestCase(false)]
        public void PivotFieldOptionsSaveTest(bool withDefaults)
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\PivotTables\PivotTables.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
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

                using (var ms = new MemoryStream())
                {
                    wb.SaveAs(ms, true);

                    ms.Position = 0;

                    using (var wbassert = new XLWorkbook(ms))
                    {
                        var wsassert = wbassert.Worksheet("pvtFieldOptionsTest");
                        var ptassert = wsassert.PivotTable("pvtFieldOptionsTest");
                        var pfassert = ptassert.RowLabels.Get("Name");
                        Assert.AreNotEqual(null, pfassert, "name save failure");
                        Assert.AreEqual("Test caption", pfassert.SubtotalCaption, "SubtotalCaption save failure");
                        Assert.AreEqual("Test name", pfassert.CustomName, "CustomName save failure");
                        AssertFieldOptions(pfassert, withDefaults);
                    }
                }
            }
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
        public void BlankPivotTableField()
        {
            using (var ms = new MemoryStream())
            {
                TestHelper.CreateAndCompare(() =>
                {
                    // Based on .\ClosedXML\ClosedXML_Examples\PivotTables\PivotTables.cs
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
        }

        [Test]
        public void SourceSheetWithWhitespace()
        {
            using (var ms = new MemoryStream())
            {
                TestHelper.CreateAndCompare(() =>
                {
                    // Based on .\ClosedXML\ClosedXML_Examples\PivotTables\PivotTables.cs
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

                    var sheet = wb.Worksheets.Add("Pastry Sales Data");
                    // Insert our list of pastry data into the "PastrySalesData" sheet at cell 1,1
                    var table = sheet.Cell(1, 1).InsertTable(pastries, "PastrySalesData", true);
                    sheet.Cell("F11").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    sheet.Columns().AdjustToContents();

                    IXLWorksheet ptSheet;
                    IXLPivotTable pt;

                    // Add a new sheet for our pivot table
                    ptSheet = wb.Worksheets.Add("pvt");

                    // Create the pivot table, using the data from the "PastrySalesData" table
                    pt = ptSheet.PivotTables.Add("pvt", ptSheet.Cell(1, 1), table.AsRange());
                    pt.ColumnLabels.Add("Name");
                    pt.RowLabels.Add("Month");

                    // The values in our table will come from the "NumberOfOrders" field
                    // The default calculation setting is a total of each row/column
                    pt.Values.Add("NumberOfOrders", "NumberOfOrdersPercentageOfBearclaw")
                        .ShowAsPercentageFrom("Name").And("Bearclaw")
                        .NumberFormat.Format = "0%";

                    ptSheet.Columns().AdjustToContents();

                    return wb;
                }, @"Other\PivotTableReferenceFiles\SourceSheetWithWhitespace\outputfile.xlsx");
            }
        }

        [Test]
        public void PivotTableWithNoneTheme()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\PivotTableReferenceFiles\PivotTableWithNoneTheme\inputfile.xlsx")))
            using (var ms = new MemoryStream())
            {
                TestHelper.CreateAndCompare(() =>
                {
                    var wb = new XLWorkbook(stream);
                    wb.SaveAs(ms);
                    return wb;
                }, @"Other\PivotTableReferenceFiles\PivotTableWithNoneTheme\outputfile.xlsx");
            }
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

                    Assert.AreEqual("Month", pageFields[0].SourceName);
                    Assert.AreEqual("Name", pageFields[1].SourceName);
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

                    Assert.AreEqual("Month", columnLabels[0].SourceName);
                    Assert.AreEqual("Name", columnLabels[1].SourceName);
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
                    pt.RowLabels.Add(XLConstants.PivotTableValuesSentinalLabel);

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

                    Assert.AreEqual("Month", rowLabels[0].SourceName);
                    Assert.AreEqual("Name", rowLabels[1].SourceName);
                    Assert.AreEqual("{{Values}}", rowLabels[2].SourceName);
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

            using (var ms = new MemoryStream())
            {
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
                    Assert.AreEqual(1, wb.Worksheets.SelectMany(ws => ws.PivotTables).Count());
                }
            }
        }

        [Test]
        public void ClearPivotTableTenderedTange()
        {
            // https://github.com/ClosedXML/ClosedXML/pull/856
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\PivotTableReferenceFiles\ClearPivotTableRenderedRangeWhenLoading\inputfile.xlsx")))
            using (var ms = new MemoryStream())
            {
                using (var wb = new XLWorkbook(stream))
                {
                    var ws = wb.Worksheet("Sheet1");
                    Assert.IsTrue(ws.Cell("B1").IsEmpty());
                    Assert.IsTrue(ws.Cell("C2").IsEmpty());
                    Assert.IsTrue(ws.Cell("D5").IsEmpty());
                    wb.SaveAs(ms);
                }

                ms.Seek(0, SeekOrigin.Begin);

                using (var wb = new XLWorkbook(ms))
                {
                    var ws = wb.Worksheet("Sheet1");
                    Assert.IsTrue(ws.Cell("B1").IsEmpty());
                    Assert.IsTrue(ws.Cell("C2").IsEmpty());
                    Assert.IsTrue(ws.Cell("D5").IsEmpty());
                }
            }
        }

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
            Assert.AreEqual(!withDefaults, field.SubtotalsAtTop, "SubtotalsAtTop save failure");
            Assert.AreEqual(!withDefaults, field.ShowBlankItems, "ShowBlankItems save failure");
            Assert.AreEqual(!withDefaults, field.Outline, "Outline save failure");
            Assert.AreEqual(!withDefaults, field.Compact, "Compact save failure");
            Assert.AreEqual(withDefaults, field.Collapsed, "Collapsed save failure");
            Assert.AreEqual(withDefaults, field.InsertBlankLines, "InsertBlankLines save failure");
            Assert.AreEqual(withDefaults, field.RepeatItemLabels, "RepeatItemLabels save failure");
            Assert.AreEqual(withDefaults, field.InsertPageBreaks, "InsertPageBreaks save failure");
            Assert.AreEqual(withDefaults, field.IncludeNewItemsInFilter, "IncludeNewItemsInFilter save failure");
        }
    }
}
