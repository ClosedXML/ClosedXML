using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLPivotTable: IXLPivotTable
    {

        public XLPivotTable()
        {
            Fields = new XLPivotFields();
            ReportFilters = new XLPivotFields();
            ColumnLabels=new XLPivotFields();
            RowLabels = new XLPivotFields();
            Values = new XLPivotValues(this);
            Theme = XLPivotTableTheme.PivotStyleLight16;

            SetExcelDefaults();
        }

        public IXLCell TargetCell { get; set; }
        public IXLRange SourceRange { get; set; }

        public IXLPivotFields Fields { get; private set; }
        public IXLPivotFields ReportFilters { get; private set; }
        public IXLPivotFields ColumnLabels { get; private set; }
        public IXLPivotFields RowLabels { get; private set; }
        public IXLPivotValues Values { get; private set; }

        public XLPivotTableTheme Theme { get; set; }		public IXLPivotTable SetTheme(XLPivotTableTheme value) { Theme = value; return this; }
        public String Name { get; set; }		public IXLPivotTable SetName(String value) { Name = value; return this; }
        public String Title { get; set; }		public IXLPivotTable SetTitle(String value) { Title = value; return this; }
        public String Description { get; set; }		public IXLPivotTable SetDescription(String value) { Description = value; return this; }

        public Boolean MergeAndCenterWithLabels { get; set; }	public IXLPivotTable SetMergeAndCenterWithLabels() { MergeAndCenterWithLabels = true; return this; }	public IXLPivotTable SetMergeAndCenterWithLabels(Boolean value) { MergeAndCenterWithLabels = value; return this; }
        public Int32 RowLabelIndent { get; set; }		public IXLPivotTable SetRowLabelIndent(Int32 value) { RowLabelIndent = value; return this; }
        public XLFilterAreaOrder FilterAreaOrder { get; set; }		public IXLPivotTable SetFilterAreaOrder(XLFilterAreaOrder value) { FilterAreaOrder = value; return this; }
        public Int32 FilterFieldsPageWrap { get; set; }		public IXLPivotTable SetFilterFieldsPageWrap(Int32 value) { FilterFieldsPageWrap = value; return this; }
        public String ErrorValueReplacement { get; set; }		public IXLPivotTable SetErrorValueReplacement(String value) { ErrorValueReplacement = value; return this; }
        public String EmptyCellReplacement { get; set; }		public IXLPivotTable SetEmptyCellReplacement(String value) { EmptyCellReplacement = value; return this; }
        public Boolean AutofitColumns { get; set; }	public IXLPivotTable SetAutofitColumns() { AutofitColumns = true; return this; }	public IXLPivotTable SetAutofitColumns(Boolean value) { AutofitColumns = value; return this; }
        public Boolean PreserveCellFormatting { get; set; }	public IXLPivotTable SetPreserveCellFormatting() { PreserveCellFormatting = true; return this; }	public IXLPivotTable SetPreserveCellFormatting(Boolean value) { PreserveCellFormatting = value; return this; }

        public Boolean ShowGrandTotalsRows { get; set; }	public IXLPivotTable SetShowGrandTotalsRows() { ShowGrandTotalsRows = true; return this; }	public IXLPivotTable SetShowGrandTotalsRows(Boolean value) { ShowGrandTotalsRows = value; return this; }
        public Boolean ShowGrandTotalsColumns { get; set; }	public IXLPivotTable SetShowGrandTotalsColumns() { ShowGrandTotalsColumns = true; return this; }	public IXLPivotTable SetShowGrandTotalsColumns(Boolean value) { ShowGrandTotalsColumns = value; return this; }
        public Boolean FilteredItemsInSubtotals { get; set; }	public IXLPivotTable SetFilteredItemsInSubtotals() { FilteredItemsInSubtotals = true; return this; }	public IXLPivotTable SetFilteredItemsInSubtotals(Boolean value) { FilteredItemsInSubtotals = value; return this; }
        public Boolean AllowMultipleFilters { get; set; }	public IXLPivotTable SetAllowMultipleFilters() { AllowMultipleFilters = true; return this; }	public IXLPivotTable SetAllowMultipleFilters(Boolean value) { AllowMultipleFilters = value; return this; }
        public Boolean UseCustomListsForSorting { get; set; }	public IXLPivotTable SetUseCustomListsForSorting() { UseCustomListsForSorting = true; return this; }	public IXLPivotTable SetUseCustomListsForSorting(Boolean value) { UseCustomListsForSorting = value; return this; }

        public Boolean ShowExpandCollapseButtons { get; set; }	public IXLPivotTable SetShowExpandCollapseButtons() { ShowExpandCollapseButtons = true; return this; }	public IXLPivotTable SetShowExpandCollapseButtons(Boolean value) { ShowExpandCollapseButtons = value; return this; }
        public Boolean ShowContextualTooltips { get; set; }	public IXLPivotTable SetShowContextualTooltips() { ShowContextualTooltips = true; return this; }	public IXLPivotTable SetShowContextualTooltips(Boolean value) { ShowContextualTooltips = value; return this; }
        public Boolean ShowPropertiesInTooltips { get; set; }	public IXLPivotTable SetShowPropertiesInTooltips() { ShowPropertiesInTooltips = true; return this; }	public IXLPivotTable SetShowPropertiesInTooltips(Boolean value) { ShowPropertiesInTooltips = value; return this; }
        public Boolean DisplayCaptionsAndDropdowns { get; set; }	public IXLPivotTable SetDisplayCaptionsAndDropdowns() { DisplayCaptionsAndDropdowns = true; return this; }	public IXLPivotTable SetDisplayCaptionsAndDropdowns(Boolean value) { DisplayCaptionsAndDropdowns = value; return this; }
        public Boolean ClassicPivotTableLayout { get; set; }	public IXLPivotTable SetClassicPivotTableLayout() { ClassicPivotTableLayout = true; return this; }	public IXLPivotTable SetClassicPivotTableLayout(Boolean value) { ClassicPivotTableLayout = value; return this; }
        public Boolean ShowValuesRow { get; set; }	public IXLPivotTable SetShowValuesRow() { ShowValuesRow = true; return this; }	public IXLPivotTable SetShowValuesRow(Boolean value) { ShowValuesRow = value; return this; }
        public Boolean ShowEmptyItemsOnRows { get; set; }	public IXLPivotTable SetShowEmptyItemsOnRows() { ShowEmptyItemsOnRows = true; return this; }	public IXLPivotTable SetShowEmptyItemsOnRows(Boolean value) { ShowEmptyItemsOnRows = value; return this; }
        public Boolean ShowEmptyItemsOnColumns { get; set; }	public IXLPivotTable SetShowEmptyItemsOnColumns() { ShowEmptyItemsOnColumns = true; return this; }	public IXLPivotTable SetShowEmptyItemsOnColumns(Boolean value) { ShowEmptyItemsOnColumns = value; return this; }
        public Boolean DisplayItemLabels { get; set; }	public IXLPivotTable SetDisplayItemLabels() { DisplayItemLabels = true; return this; }	public IXLPivotTable SetDisplayItemLabels(Boolean value) { DisplayItemLabels = value; return this; }
        public Boolean SortFieldsAtoZ { get; set; }	public IXLPivotTable SetSortFieldsAtoZ() { SortFieldsAtoZ = true; return this; }	public IXLPivotTable SetSortFieldsAtoZ(Boolean value) { SortFieldsAtoZ = value; return this; }

        public Boolean PrintExpandCollapsedButtons { get; set; }	public IXLPivotTable SetPrintExpandCollapsedButtons() { PrintExpandCollapsedButtons = true; return this; }	public IXLPivotTable SetPrintExpandCollapsedButtons(Boolean value) { PrintExpandCollapsedButtons = value; return this; }
        public Boolean RepeatRowLabels { get; set; }	public IXLPivotTable SetRepeatRowLabels() { RepeatRowLabels = true; return this; }	public IXLPivotTable SetRepeatRowLabels(Boolean value) { RepeatRowLabels = value; return this; }
        public Boolean PrintTitles { get; set; }	public IXLPivotTable SetPrintTitles() { PrintTitles = true; return this; }	public IXLPivotTable SetPrintTitles(Boolean value) { PrintTitles = value; return this; }

        public Boolean SaveSourceData { get; set; }	public IXLPivotTable SetSaveSourceData() { SaveSourceData = true; return this; }	public IXLPivotTable SetSaveSourceData(Boolean value) { SaveSourceData = value; return this; }
        public Boolean EnableShowDetails { get; set; }	public IXLPivotTable SetEnableShowDetails() { EnableShowDetails = true; return this; }	public IXLPivotTable SetEnableShowDetails(Boolean value) { EnableShowDetails = value; return this; }
        public Boolean RefreshDataOnOpen { get; set; }	public IXLPivotTable SetRefreshDataOnOpen() { RefreshDataOnOpen = true; return this; }	public IXLPivotTable SetRefreshDataOnOpen(Boolean value) { RefreshDataOnOpen = value; return this; }
        public XLItemsToRetain ItemsToRetainPerField { get; set; }		public IXLPivotTable SetItemsToRetainPerField(XLItemsToRetain value) { ItemsToRetainPerField = value; return this; }
        public Boolean EnableCellEditing { get; set; }	public IXLPivotTable SetEnableCellEditing() { EnableCellEditing = true; return this; }	public IXLPivotTable SetEnableCellEditing(Boolean value) { EnableCellEditing = value; return this; }


        public Boolean ShowRowHeaders { get; set; }	public IXLPivotTable SetShowRowHeaders() { ShowRowHeaders = true; return this; }	public IXLPivotTable SetShowRowHeaders(Boolean value) { ShowRowHeaders = value; return this; }
        public Boolean ShowColumnHeaders { get; set; }	public IXLPivotTable SetShowColumnHeaders() { ShowColumnHeaders = true; return this; }	public IXLPivotTable SetShowColumnHeaders(Boolean value) { ShowColumnHeaders = value; return this; }
        public Boolean ShowRowStripes { get; set; }	public IXLPivotTable SetShowRowStripes() { ShowRowStripes = true; return this; }	public IXLPivotTable SetShowRowStripes(Boolean value) { ShowRowStripes = value; return this; }
        public Boolean ShowColumnStripes { get; set; }	public IXLPivotTable SetShowColumnStripes() { ShowColumnStripes = true; return this; }	public IXLPivotTable SetShowColumnStripes(Boolean value) { ShowColumnStripes = value; return this; }
        public XLPivotSubtotals Subtotals { get; set; }		public IXLPivotTable SetSubtotals(XLPivotSubtotals value) { Subtotals = value; return this; }

        public XLPivotLayout Layout
        {
            set { Fields.ForEach(f=>f.SetLayout(value)); }
        }

        public IXLPivotTable SetLayout(XLPivotLayout value) { Layout = value; return this; }

        public Boolean InsertBlankLines
        {
            set { Fields.ForEach(f=>f.SetInsertBlankLines(value)); }
        }

        public IXLPivotTable SetInsertBlankLines() { InsertBlankLines = true; return this; }	public IXLPivotTable SetInsertBlankLines(Boolean value) { InsertBlankLines = value; return this; }

        internal String RelId { get; set; }
        internal String CacheDefinitionRelId { get; set; }
        internal String WorkbookCacheRelId { get; set; }

        private void SetExcelDefaults()
        {
            EmptyCellReplacement = String.Empty;
            AutofitColumns = true;
            PreserveCellFormatting = true;
            ShowGrandTotalsColumns = true;
            ShowGrandTotalsRows = true;
            UseCustomListsForSorting = true;
            ShowExpandCollapseButtons = true;
            ShowContextualTooltips = true;
            DisplayCaptionsAndDropdowns = true;
            RepeatRowLabels = true;
            SaveSourceData = true;
            EnableShowDetails = true;
            ShowColumnHeaders = true;
            ShowRowHeaders = true;
        }

    }
}
