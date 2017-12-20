using ClosedXML.Excel;
using NUnit.Framework;
using System.IO;

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

                var range = table.DataRange;
                var header = ws.Range(1, 1, 1, 3);
                var dataRange = ws.Range(header.FirstCell(), range.LastCell());

                var ptSheet = wb.Worksheets.Add("BlankPivotTable");
                var pt = ptSheet.PivotTables.AddNew("pvt", ptSheet.Cell(1, 1), dataRange);

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

                var range = table.DataRange;
                var header = ws.Range(1, 1, 1, 3);
                var dataRange = ws.Range(header.FirstCell(), range.LastCell());

                var ptSheet = wb.Worksheets.Add("BlankPivotTable");
                var pt = ptSheet.PivotTables.AddNew("pvtOptionsTest", ptSheet.Cell(1, 1), dataRange);

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
    }
}
