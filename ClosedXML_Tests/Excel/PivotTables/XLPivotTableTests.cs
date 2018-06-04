using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

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

        [Test]
        public void PivotFieldFormatsTest()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\PivotTables\PivotTables.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheet("PastrySalesData");
                var table = ws.Table("PastrySalesData");
                var ptSheet = wb.Worksheets.Add("PivotFieldFormats");
                var pt = ptSheet.PivotTables.Add("pvtFormats", ptSheet.Cell(1, 1), table);
                pt.SetSubtotals(XLPivotSubtotals.AtBottom);
                var clmn = pt.ColumnLabels.Add("Month");
                var row = pt.RowLabels.Add("Name")
                    .SetSubtotalCaption("Test caption")
                    .SetCustomName("Test name")
                    .AddSubtotal(XLSubtotalFunction.Sum);

                var numberOfOrders = pt.Values.Add("NumberOfOrders").SetSummaryFormula(XLPivotSummary.Sum);
                var quality = pt.Values.Add("Quality").SetSummaryFormula(XLPivotSummary.Sum);

                pt.PivotFormats.Add(b => { b.ForSubtotal(row).Style.Fill.BackgroundColor = XLColor.Blue; });
                pt.PivotFormats.Add(b => { b.ForLabel(clmn).Style.Fill.BackgroundColor = XLColor.Amber; });
                pt.PivotFormats.Add(b => { b.ForHeader(clmn).Style.Font.FontColor = XLColor.Yellow; });
                pt.PivotFormats.Add(b => { b.ForGrandRow().Style.Font.FontColor = XLColor.VenetianRed; });
                pt.PivotFormats.Add(b => {
                    b.ForData(row)
                    .AndWith(clmn, v=>v.ToString() == "May")
                    .ForValueField(numberOfOrders)
                    .Style.Font.FontColor = XLColor.Green;
                });

                using (var ms = new MemoryStream())
                {
                    wb.SaveAs(ms, true);

                    ms.Position = 0;

                    using (var wbassert = new XLWorkbook(ms))
                    {
                        var wsassert = wbassert.Worksheet("PivotFieldFormats");
                        var ptassert = wsassert.PivotTable("pvtFormats");
                        var formats = ptassert.PivotFormats.ToList();
                        Assert.AreEqual(5, formats.Count, "PivotFormats, loaded count must be 5");

                        // ForSubtotal
                        Assert.AreEqual("Name", formats[0].FieldReferences.First().FieldName, "ForSubtotal field name must be 'Name'");
                        Assert.AreEqual(true, formats[0].FieldReferences.First().DefaultSubtotal, "ForSubtotal field DefaultSubtotal must be true");

                        // ForLabel
                        Assert.AreEqual("Month", formats[1].FieldReferences.First().FieldName, "ForLabel field name must be 'Month'");
                        Assert.AreEqual(-1, formats[1].FieldReferences.First().Value, "ForLabel field Value must be -1");
                        Assert.AreEqual(false, formats[1].DataOnly, "ForLabel field DataOnly must be false");
                        Assert.AreEqual(true, formats[1].LabelOnly, "ForLabel field LabelOnly must be true");

                        // ForHeader
                        Assert.AreEqual("Month", formats[2].FieldName, "ForHeader field name must be 'Month'");
                        Assert.AreEqual(false, formats[2].DataOnly, "ForHeader field DataOnly must be false");
                        Assert.AreEqual(true, formats[2].LabelOnly, "ForHeader field LabelOnly must be true");
                        Assert.AreEqual(false, formats[2].Outline, "ForHeader field Outline must be false");
                        Assert.AreEqual(PivotAreaValues.Button, formats[2].AreaType, "ForHeader field AreaType must be false");
                        Assert.AreEqual(PivotTableAxisValues.AxisColumn, formats[2].Axis, "ForHeader field Axis must be false");

                        // ForGrandRow
                        Assert.AreEqual(PivotTableAxisValues.AxisRow, formats[3].Axis, "ForGrandRow field Axis must be AxisRow");
                        Assert.AreEqual(true, formats[3].GrandRow, "ForGrandRow field GrandRow must be true");

                        // ForData
                        var refs = formats[4].FieldReferences.ToList();
                        Assert.AreEqual(3, refs.Count, "ForData FieldReferences.Count must be 3");
                        Assert.IsTrue(refs.Any(i=>i.FieldName == "Month"), "ForData FieldReferences must contains 'Month'");
                        Assert.IsTrue(refs.Any(i=>i.FieldName == "Name"), "ForData FieldReferences must contains 'Name'");
                        Assert.IsTrue(refs.Any(i=>i.FieldName == XLConstants.PivotTableValuesSentinalLabel), "ForData FieldReferences must contains '{{Values}}'");
                        Assert.AreEqual(1, refs.First(i=>i.FieldName == "Month").Value, "ForData FieldReferences value must be 1 (May)");
                        Assert.AreEqual(0, refs.First(i=>i.FieldName == XLConstants.PivotTableValuesSentinalLabel).Value, "ForData FieldReferences value must be 0");
                        Assert.IsTrue(formats[4].DataOnly, "ForData field DataOnly must be true");
                        Assert.IsFalse(formats[4].LabelOnly, "ForData field LabelOnly must be false");
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
