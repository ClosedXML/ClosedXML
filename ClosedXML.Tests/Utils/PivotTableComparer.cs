using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Tests
{
    internal class PivotTableComparer : IEqualityComparer<XLPivotTable>
    {
        private readonly bool _compareName;
        private readonly bool _compareRelId;
        private readonly bool _compareTargetCellAddress;

        public PivotTableComparer()
            : this(compareName: true, compareRelId: false, compareTargetCellAddress: true)
        { }

        public PivotTableComparer(bool compareName, bool compareRelId, bool compareTargetCellAddress)
        {
            this._compareName = compareName;
            this._compareRelId = compareRelId;
            this._compareTargetCellAddress = compareTargetCellAddress;
        }

        public bool Equals(XLPivotTable x, XLPivotTable y)
        {
            if (x == null && y == null) return true;

            if (x == null || y == null) return false;

            return
                (!_compareName || StringComparer.CurrentCulture.Equals(x.Name, y.Name))
                && (!_compareRelId || StringComparer.CurrentCulture.Equals(x.RelId, y.RelId))

                && x.ReportFilters.Count().Equals(y.ReportFilters.Count())
                && x.ColumnLabels.Count().Equals(y.ColumnLabels.Count())
                && x.RowLabels.Count().Equals(y.RowLabels.Count())
                && x.Values.Count().Equals(y.Values.Count())
                && (!_compareTargetCellAddress || x.TargetCell.Address.ColumnLetter.Equals(y.TargetCell.Address.ColumnLetter))
                && (!_compareTargetCellAddress || x.TargetCell.Address.RowNumber.Equals(y.TargetCell.Address.RowNumber))

                && StringComparer.CurrentCulture.Equals(x.Title, y.Title)
                && StringComparer.CurrentCulture.Equals(x.Description, y.Description)
                && StringComparer.CurrentCulture.Equals(x.ColumnHeaderCaption, y.ColumnHeaderCaption)
                && StringComparer.CurrentCulture.Equals(x.RowHeaderCaption, y.RowHeaderCaption)
                && x.MergeAndCenterWithLabels.Equals(y.MergeAndCenterWithLabels)
                && x.RowLabelIndent.Equals(y.RowLabelIndent)
                && x.FilterAreaOrder.Equals(y.FilterAreaOrder)
                && x.FilterFieldsPageWrap.Equals(y.FilterFieldsPageWrap)
                && StringComparer.CurrentCulture.Equals(x.ErrorValueReplacement, y.ErrorValueReplacement)
                && StringComparer.CurrentCulture.Equals(x.EmptyCellReplacement, y.EmptyCellReplacement)
                && x.AutofitColumns.Equals(y.AutofitColumns)
                && x.PreserveCellFormatting.Equals(y.PreserveCellFormatting)
                && x.ShowGrandTotalsColumns.Equals(y.ShowGrandTotalsColumns)
                && x.ShowGrandTotalsRows.Equals(y.ShowGrandTotalsRows)
                && x.FilteredItemsInSubtotals.Equals(y.FilteredItemsInSubtotals)
                && x.AllowMultipleFilters.Equals(y.AllowMultipleFilters)
                && x.UseCustomListsForSorting.Equals(y.UseCustomListsForSorting)
                && x.ShowExpandCollapseButtons.Equals(y.ShowExpandCollapseButtons)
                && x.ShowContextualTooltips.Equals(y.ShowContextualTooltips)
                && x.ShowPropertiesInTooltips.Equals(y.ShowPropertiesInTooltips)
                && x.DisplayCaptionsAndDropdowns.Equals(y.DisplayCaptionsAndDropdowns)
                && x.ClassicPivotTableLayout.Equals(y.ClassicPivotTableLayout)
                && x.ShowValuesRow.Equals(y.ShowValuesRow)
                && x.ShowEmptyItemsOnColumns.Equals(y.ShowEmptyItemsOnColumns)
                && x.ShowEmptyItemsOnRows.Equals(y.ShowEmptyItemsOnRows)
                && x.DisplayItemLabels.Equals(y.DisplayItemLabels)
                && x.SortFieldsAtoZ.Equals(y.SortFieldsAtoZ)
                && x.PrintExpandCollapsedButtons.Equals(y.PrintExpandCollapsedButtons)
                && x.RepeatRowLabels.Equals(y.RepeatRowLabels)
                && x.PrintTitles.Equals(y.PrintTitles)
                && x.SaveSourceData.Equals(y.SaveSourceData)
                && x.EnableShowDetails.Equals(y.EnableShowDetails)
                && x.RefreshDataOnOpen.Equals(y.RefreshDataOnOpen)
                && x.ItemsToRetainPerField.Equals(y.ItemsToRetainPerField)
                && x.EnableCellEditing.Equals(y.EnableCellEditing)
                && x.ShowRowHeaders.Equals(y.ShowRowHeaders)
                && x.ShowColumnHeaders.Equals(y.ShowColumnHeaders)
                && x.ShowRowStripes.Equals(y.ShowRowStripes)
                && x.ShowColumnStripes.Equals(y.ShowColumnStripes)
                && x.Theme.Equals(y.Theme);
        }

        public int GetHashCode(XLPivotTable obj)
        {
            throw new NotImplementedException();
        }
    }
}
