using ClosedXML.Excel.CalcEngine;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace ClosedXML.Excel
{
    [DebuggerDisplay("{Name}")]
    internal class XLPivotTable : IXLPivotTable
    {
        private String _name;
        public Guid Guid { get; private set; }

        public XLPivotTable(IXLWorksheet worksheet)
        {
            this.Worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
            this.Guid = Guid.NewGuid();

            Fields = new XLPivotFields(this);
            ReportFilters = new XLPivotFields(this);
            ColumnLabels = new XLPivotFields(this);
            RowLabels = new XLPivotFields(this);
            Values = new XLPivotValues(this);
            Theme = XLPivotTableTheme.PivotStyleLight16;

            SetExcelDefaults();
        }

        public IXLCell TargetCell { get; set; }

        private IXLRange sourceRange;

        public IXLRange SourceRange
        {
            get { return sourceRange; }
            set
            {
                if (value is IXLTable)
                    SourceType = XLPivotTableSourceType.Table;
                else
                    SourceType = XLPivotTableSourceType.Range;

                sourceRange = value;
            }
        }

        public IXLTable SourceTable
        {
            get { return SourceRange as IXLTable; }
            set { SourceRange = value; }
        }

        public XLPivotTableSourceType SourceType { get; private set; }

        public IEnumerable<string> SourceRangeFieldsAvailable
        {
            get { return this.SourceRange.FirstRow().Cells().Select(c => c.GetString()); }
        }

        public IXLPivotFields Fields { get; private set; }
        public IXLPivotFields ReportFilters { get; private set; }
        public IXLPivotFields ColumnLabels { get; private set; }
        public IXLPivotFields RowLabels { get; private set; }
        public IXLPivotValues Values { get; private set; }

        public IEnumerable<IXLPivotField> ImplementedFields
        {
            get
            {
                foreach (var pf in ReportFilters)
                    yield return pf;

                foreach (var pf in RowLabels)
                    yield return pf;

                foreach (var pf in ColumnLabels)
                    yield return pf;
            }
        }

        public XLPivotTableTheme Theme { get; set; }

        public IXLPivotTable CopyTo(IXLCell targetCell)
        {
            var addressComparer = new XLAddressComparer(ignoreFixed: true);
            if (addressComparer.Equals(targetCell.Address, this.TargetCell.Address))
                throw new InvalidOperationException("Cannot copy pivot table to the target cell.");

            var targetSheet = targetCell.Worksheet;

            var pivotTableName = this.Name;

            int i = 0;
            var pivotTableNames = targetSheet.PivotTables.Select(pvt => pvt.Name).ToList();
            while (!XLHelper.ValidateName("pivot table", pivotTableName, "", pivotTableNames, out var _))
            {
                i++;
                pivotTableName = this.Name + i.ToInvariantString();
            }

            var newPivotTable = this.SourceType switch
            {
                XLPivotTableSourceType.Table => targetSheet.PivotTables.Add(pivotTableName, targetCell, this.SourceTable) as XLPivotTable,
                XLPivotTableSourceType.Range => targetSheet.PivotTables.Add(pivotTableName, targetCell, this.SourceRange) as XLPivotTable,
                _ => throw new NotImplementedException(),
            };

            newPivotTable.RelId = null;

            static void copyPivotField(IXLPivotField originalPivotField, IXLPivotField newPivotField)
            {
                newPivotField
                    .SetSort(originalPivotField.SortType)
                    .SetSubtotalCaption(originalPivotField.SubtotalCaption)
                    .SetIncludeNewItemsInFilter(originalPivotField.IncludeNewItemsInFilter)
                    .SetRepeatItemLabels(originalPivotField.RepeatItemLabels)
                    .SetInsertBlankLines(originalPivotField.InsertBlankLines)
                    .SetShowBlankItems(originalPivotField.ShowBlankItems)
                    .SetInsertPageBreaks(originalPivotField.InsertPageBreaks)
                    .SetCollapsed(originalPivotField.Collapsed);

                if (originalPivotField.SubtotalsAtTop.HasValue)
                    newPivotField.SetSubtotalsAtTop(originalPivotField.SubtotalsAtTop.Value);

                newPivotField.AddSelectedValues(originalPivotField.SelectedValues);
            }

            foreach (var rf in ReportFilters)
                copyPivotField(rf, newPivotTable.ReportFilters.Add(rf.SourceName, rf.CustomName));

            foreach (var cl in ColumnLabels)
                copyPivotField(cl, newPivotTable.ColumnLabels.Add(cl.SourceName, cl.CustomName));

            foreach (var rl in RowLabels)
                copyPivotField(rl, newPivotTable.RowLabels.Add(rl.SourceName, rl.CustomName));

            foreach (var v in Values)
            {
                var pivotValue = newPivotTable.Values.Add(v.SourceName, v.CustomName)
                    .SetSummaryFormula(v.SummaryFormula)
                    .SetCalculation(v.Calculation)
                    .SetCalculationItem(v.CalculationItem)
                    .SetBaseField(v.BaseField)
                    .SetBaseItem(v.BaseItem);

                pivotValue.NumberFormat.NumberFormatId = v.NumberFormat.NumberFormatId;
                pivotValue.NumberFormat.Format = v.NumberFormat.Format;
            }

            newPivotTable.Title = Title;
            newPivotTable.Description = Description;
            newPivotTable.ColumnHeaderCaption = ColumnHeaderCaption;
            newPivotTable.RowHeaderCaption = RowHeaderCaption;
            newPivotTable.MergeAndCenterWithLabels = MergeAndCenterWithLabels;
            newPivotTable.RowLabelIndent = RowLabelIndent;
            newPivotTable.FilterAreaOrder = FilterAreaOrder;
            newPivotTable.FilterFieldsPageWrap = FilterFieldsPageWrap;
            newPivotTable.ErrorValueReplacement = ErrorValueReplacement;
            newPivotTable.EmptyCellReplacement = EmptyCellReplacement;
            newPivotTable.AutofitColumns = AutofitColumns;
            newPivotTable.PreserveCellFormatting = PreserveCellFormatting;
            newPivotTable.ShowGrandTotalsColumns = ShowGrandTotalsColumns;
            newPivotTable.ShowGrandTotalsRows = ShowGrandTotalsRows;
            newPivotTable.FilteredItemsInSubtotals = FilteredItemsInSubtotals;
            newPivotTable.AllowMultipleFilters = AllowMultipleFilters;
            newPivotTable.UseCustomListsForSorting = UseCustomListsForSorting;
            newPivotTable.ShowExpandCollapseButtons = ShowExpandCollapseButtons;
            newPivotTable.ShowContextualTooltips = ShowContextualTooltips;
            newPivotTable.ShowPropertiesInTooltips = ShowPropertiesInTooltips;
            newPivotTable.DisplayCaptionsAndDropdowns = DisplayCaptionsAndDropdowns;
            newPivotTable.ClassicPivotTableLayout = ClassicPivotTableLayout;
            newPivotTable.ShowValuesRow = ShowValuesRow;
            newPivotTable.ShowEmptyItemsOnColumns = ShowEmptyItemsOnColumns;
            newPivotTable.ShowEmptyItemsOnRows = ShowEmptyItemsOnRows;
            newPivotTable.DisplayItemLabels = DisplayItemLabels;
            newPivotTable.SortFieldsAtoZ = SortFieldsAtoZ;
            newPivotTable.PrintExpandCollapsedButtons = PrintExpandCollapsedButtons;
            newPivotTable.RepeatRowLabels = RepeatRowLabels;
            newPivotTable.PrintTitles = PrintTitles;
            newPivotTable.SaveSourceData = SaveSourceData;
            newPivotTable.EnableShowDetails = EnableShowDetails;
            newPivotTable.RefreshDataOnOpen = RefreshDataOnOpen;
            newPivotTable.ItemsToRetainPerField = ItemsToRetainPerField;
            newPivotTable.EnableCellEditing = EnableCellEditing;
            newPivotTable.ShowRowHeaders = ShowRowHeaders;
            newPivotTable.ShowColumnHeaders = ShowColumnHeaders;
            newPivotTable.ShowRowStripes = ShowRowStripes;
            newPivotTable.ShowColumnStripes = ShowColumnStripes;
            newPivotTable.Theme = Theme;
            // TODO: Copy Styleformats

            return newPivotTable;
        }

        public IXLPivotTable SetTheme(XLPivotTableTheme value)
        {
            Theme = value; return this;
        }

        public String Name
        {
            get { return _name; }
            set
            {
                if (_name == value) return;

                var oldname = _name ?? string.Empty;

                if (!XLHelper.ValidateName("pivot table", value, oldname, Worksheet.PivotTables.Select(pvt => pvt.Name), out String message))
                    throw new ArgumentException(message, nameof(value));

                _name = value;

                if (!String.IsNullOrWhiteSpace(oldname) && !String.Equals(oldname, _name, StringComparison.OrdinalIgnoreCase))
                {
                    Worksheet.PivotTables.Delete(oldname);
                    (Worksheet.PivotTables as XLPivotTables).Add(_name, this);
                }
            }
        }

        public IXLPivotTable SetName(String value)
        {
            Name = value; return this;
        }

        public String Title { get; set; }

        public IXLPivotTable SetTitle(String value)
        {
            Title = value; return this;
        }

        public String Description { get; set; }

        public IXLPivotTable SetDescription(String value)
        {
            Description = value; return this;
        }

        public String ColumnHeaderCaption { get; set; }

        public IXLPivotTable SetColumnHeaderCaption(String value)
        {
            ColumnHeaderCaption = value;
            return this;
        }

        public String RowHeaderCaption { get; set; }

        public IXLPivotTable SetRowHeaderCaption(String value)
        {
            RowHeaderCaption = value;
            return this;
        }

        public Boolean MergeAndCenterWithLabels { get; set; }

        public IXLPivotTable SetMergeAndCenterWithLabels()
        {
            MergeAndCenterWithLabels = true; return this;
        }

        public IXLPivotTable SetMergeAndCenterWithLabels(Boolean value)
        {
            MergeAndCenterWithLabels = value; return this;
        }

        public Int32 RowLabelIndent { get; set; }

        public IXLPivotTable SetRowLabelIndent(Int32 value)
        {
            RowLabelIndent = value; return this;
        }

        public XLFilterAreaOrder FilterAreaOrder { get; set; }

        public IXLPivotTable SetFilterAreaOrder(XLFilterAreaOrder value)
        {
            FilterAreaOrder = value; return this;
        }

        public Int32 FilterFieldsPageWrap { get; set; }

        public IXLPivotTable SetFilterFieldsPageWrap(Int32 value)
        {
            FilterFieldsPageWrap = value; return this;
        }

        public String ErrorValueReplacement { get; set; }

        public IXLPivotTable SetErrorValueReplacement(String value)
        {
            ErrorValueReplacement = value; return this;
        }

        public String EmptyCellReplacement { get; set; }

        public IXLPivotTable SetEmptyCellReplacement(String value)
        {
            EmptyCellReplacement = value; return this;
        }

        public Boolean AutofitColumns { get; set; }

        public IXLPivotTable SetAutofitColumns()
        {
            AutofitColumns = true; return this;
        }

        public IXLPivotTable SetAutofitColumns(Boolean value)
        {
            AutofitColumns = value; return this;
        }

        public Boolean PreserveCellFormatting { get; set; }

        public IXLPivotTable SetPreserveCellFormatting()
        {
            PreserveCellFormatting = true; return this;
        }

        public IXLPivotTable SetPreserveCellFormatting(Boolean value)
        {
            PreserveCellFormatting = value; return this;
        }

        public Boolean ShowGrandTotalsRows { get; set; }

        public IXLPivotTable SetShowGrandTotalsRows()
        {
            ShowGrandTotalsRows = true; return this;
        }

        public IXLPivotTable SetShowGrandTotalsRows(Boolean value)
        {
            ShowGrandTotalsRows = value; return this;
        }

        public Boolean ShowGrandTotalsColumns { get; set; }

        public IXLPivotTable SetShowGrandTotalsColumns()
        {
            ShowGrandTotalsColumns = true; return this;
        }

        public IXLPivotTable SetShowGrandTotalsColumns(Boolean value)
        {
            ShowGrandTotalsColumns = value; return this;
        }

        public Boolean FilteredItemsInSubtotals { get; set; }

        public IXLPivotTable SetFilteredItemsInSubtotals()
        {
            FilteredItemsInSubtotals = true; return this;
        }

        public IXLPivotTable SetFilteredItemsInSubtotals(Boolean value)
        {
            FilteredItemsInSubtotals = value; return this;
        }

        public Boolean AllowMultipleFilters { get; set; }

        public IXLPivotTable SetAllowMultipleFilters()
        {
            AllowMultipleFilters = true; return this;
        }

        public IXLPivotTable SetAllowMultipleFilters(Boolean value)
        {
            AllowMultipleFilters = value; return this;
        }

        public Boolean UseCustomListsForSorting { get; set; }

        public IXLPivotTable SetUseCustomListsForSorting()
        {
            UseCustomListsForSorting = true; return this;
        }

        public IXLPivotTable SetUseCustomListsForSorting(Boolean value)
        {
            UseCustomListsForSorting = value; return this;
        }

        public Boolean ShowExpandCollapseButtons { get; set; }

        public IXLPivotTable SetShowExpandCollapseButtons()
        {
            ShowExpandCollapseButtons = true; return this;
        }

        public IXLPivotTable SetShowExpandCollapseButtons(Boolean value)
        {
            ShowExpandCollapseButtons = value; return this;
        }

        public Boolean ShowContextualTooltips { get; set; }

        public IXLPivotTable SetShowContextualTooltips()
        {
            ShowContextualTooltips = true; return this;
        }

        public IXLPivotTable SetShowContextualTooltips(Boolean value)
        {
            ShowContextualTooltips = value; return this;
        }

        public Boolean ShowPropertiesInTooltips { get; set; }

        public IXLPivotTable SetShowPropertiesInTooltips()
        {
            ShowPropertiesInTooltips = true; return this;
        }

        public IXLPivotTable SetShowPropertiesInTooltips(Boolean value)
        {
            ShowPropertiesInTooltips = value; return this;
        }

        public Boolean DisplayCaptionsAndDropdowns { get; set; }

        public IXLPivotTable SetDisplayCaptionsAndDropdowns()
        {
            DisplayCaptionsAndDropdowns = true; return this;
        }

        public IXLPivotTable SetDisplayCaptionsAndDropdowns(Boolean value)
        {
            DisplayCaptionsAndDropdowns = value; return this;
        }

        public Boolean ClassicPivotTableLayout { get; set; }

        public IXLPivotTable SetClassicPivotTableLayout()
        {
            ClassicPivotTableLayout = true; return this;
        }

        public IXLPivotTable SetClassicPivotTableLayout(Boolean value)
        {
            ClassicPivotTableLayout = value; return this;
        }

        public Boolean ShowValuesRow { get; set; }

        public IXLPivotTable SetShowValuesRow()
        {
            ShowValuesRow = true; return this;
        }

        public IXLPivotTable SetShowValuesRow(Boolean value)
        {
            ShowValuesRow = value; return this;
        }

        public Boolean ShowEmptyItemsOnRows { get; set; }

        public IXLPivotTable SetShowEmptyItemsOnRows()
        {
            ShowEmptyItemsOnRows = true; return this;
        }

        public IXLPivotTable SetShowEmptyItemsOnRows(Boolean value)
        {
            ShowEmptyItemsOnRows = value; return this;
        }

        public Boolean ShowEmptyItemsOnColumns { get; set; }

        public IXLPivotTable SetShowEmptyItemsOnColumns()
        {
            ShowEmptyItemsOnColumns = true; return this;
        }

        public IXLPivotTable SetShowEmptyItemsOnColumns(Boolean value)
        {
            ShowEmptyItemsOnColumns = value; return this;
        }

        public Boolean DisplayItemLabels { get; set; }

        public IXLPivotTable SetDisplayItemLabels()
        {
            DisplayItemLabels = true; return this;
        }

        public IXLPivotTable SetDisplayItemLabels(Boolean value)
        {
            DisplayItemLabels = value; return this;
        }

        public Boolean SortFieldsAtoZ { get; set; }

        public IXLPivotTable SetSortFieldsAtoZ()
        {
            SortFieldsAtoZ = true; return this;
        }

        public IXLPivotTable SetSortFieldsAtoZ(Boolean value)
        {
            SortFieldsAtoZ = value; return this;
        }

        public Boolean PrintExpandCollapsedButtons { get; set; }

        public IXLPivotTable SetPrintExpandCollapsedButtons()
        {
            PrintExpandCollapsedButtons = true; return this;
        }

        public IXLPivotTable SetPrintExpandCollapsedButtons(Boolean value)
        {
            PrintExpandCollapsedButtons = value; return this;
        }

        public Boolean RepeatRowLabels { get; set; }

        public IXLPivotTable SetRepeatRowLabels()
        {
            RepeatRowLabels = true; return this;
        }

        public IXLPivotTable SetRepeatRowLabels(Boolean value)
        {
            RepeatRowLabels = value; return this;
        }

        public Boolean PrintTitles { get; set; }

        public IXLPivotTable SetPrintTitles()
        {
            PrintTitles = true; return this;
        }

        public IXLPivotTable SetPrintTitles(Boolean value)
        {
            PrintTitles = value; return this;
        }

        public Boolean SaveSourceData { get; set; }

        public IXLPivotTable SetSaveSourceData()
        {
            SaveSourceData = true; return this;
        }

        public IXLPivotTable SetSaveSourceData(Boolean value)
        {
            SaveSourceData = value; return this;
        }

        public Boolean EnableShowDetails { get; set; }

        public IXLPivotTable SetEnableShowDetails()
        {
            EnableShowDetails = true; return this;
        }

        public IXLPivotTable SetEnableShowDetails(Boolean value)
        {
            EnableShowDetails = value; return this;
        }

        public Boolean RefreshDataOnOpen { get; set; }

        public IXLPivotTable SetRefreshDataOnOpen()
        {
            RefreshDataOnOpen = true; return this;
        }

        public IXLPivotTable SetRefreshDataOnOpen(Boolean value)
        {
            RefreshDataOnOpen = value; return this;
        }

        public XLItemsToRetain ItemsToRetainPerField { get; set; }

        public IXLPivotTable SetItemsToRetainPerField(XLItemsToRetain value)
        {
            ItemsToRetainPerField = value; return this;
        }

        public Boolean EnableCellEditing { get; set; }

        public IXLPivotTable SetEnableCellEditing()
        {
            EnableCellEditing = true; return this;
        }

        public IXLPivotTable SetEnableCellEditing(Boolean value)
        {
            EnableCellEditing = value; return this;
        }

        public Boolean ShowRowHeaders { get; set; }

        public IXLPivotTable SetShowRowHeaders()
        {
            ShowRowHeaders = true; return this;
        }

        public IXLPivotTable SetShowRowHeaders(Boolean value)
        {
            ShowRowHeaders = value; return this;
        }

        public Boolean ShowColumnHeaders { get; set; }

        public IXLPivotTable SetShowColumnHeaders()
        {
            ShowColumnHeaders = true; return this;
        }

        public IXLPivotTable SetShowColumnHeaders(Boolean value)
        {
            ShowColumnHeaders = value; return this;
        }

        public Boolean ShowRowStripes { get; set; }

        public IXLPivotTable SetShowRowStripes()
        {
            ShowRowStripes = true; return this;
        }

        public IXLPivotTable SetShowRowStripes(Boolean value)
        {
            ShowRowStripes = value; return this;
        }

        public Boolean ShowColumnStripes { get; set; }

        public IXLPivotTable SetShowColumnStripes()
        {
            ShowColumnStripes = true; return this;
        }

        public IXLPivotTable SetShowColumnStripes(Boolean value)
        {
            ShowColumnStripes = value; return this;
        }

        public XLPivotSubtotals Subtotals { get; set; }

        public IXLPivotTable SetSubtotals(XLPivotSubtotals value)
        {
            Subtotals = value; return this;
        }

        public XLPivotLayout Layout
        {
            set { Fields.ForEach(f => f.SetLayout(value)); }
        }

        public IXLPivotTable SetLayout(XLPivotLayout value)
        {
            Layout = value; return this;
        }

        public Boolean InsertBlankLines
        {
            set { Fields.ForEach(f => f.SetInsertBlankLines(value)); }
        }

        public IXLPivotTable SetInsertBlankLines()
        {
            InsertBlankLines = true; return this;
        }

        public IXLPivotTable SetInsertBlankLines(Boolean value)
        {
            InsertBlankLines = value; return this;
        }

        internal String RelId { get; set; }
        internal String CacheDefinitionRelId { get; set; }
        internal String WorkbookCacheRelId { get; set; }

        private void SetExcelDefaults()
        {
            EmptyCellReplacement = String.Empty;
            SaveSourceData = true;
            ShowColumnHeaders = true;
            ShowRowHeaders = true;

            // source http://www.datypic.com/sc/ooxml/e-ssml_pivotTableDefinition.html
            DisplayItemLabels = true;	//	Show Item Names
            ShowExpandCollapseButtons = true;	//	Show Expand Collapse
            PrintExpandCollapsedButtons = false;	//	Print Drill Indicators
            ShowPropertiesInTooltips = true;	//	Show Member Property ToolTips
            ShowContextualTooltips = true;	//	Show ToolTips on Data
            EnableShowDetails = true;	//	Enable Drill Down
            PreserveCellFormatting = true;	//	Preserve Formatting
            AutofitColumns = false;	//	Auto Formatting
            FilterAreaOrder = XLFilterAreaOrder.DownThenOver;	//	Page Over Then Down
            FilteredItemsInSubtotals = false;	//	Subtotal Hidden Items
            ShowGrandTotalsRows = true;	//	Row Grand Totals
            ShowGrandTotalsColumns = true;	//	Grand Totals On Columns
            PrintTitles = false;	//	Field Print Titles
            RepeatRowLabels = false;	//	Item Print Titles
            MergeAndCenterWithLabels = false;	//	Merge Titles
            RowLabelIndent = 1;	//	Indentation for Compact Axis
            ShowEmptyItemsOnRows = false;	//	Show Empty Row
            ShowEmptyItemsOnColumns = false;	//	Show Empty Column
            DisplayCaptionsAndDropdowns = true;	//	Show Field Headers
            ClassicPivotTableLayout = false;	//	Enable Drop Zones
            AllowMultipleFilters = true;	//	Multiple Field Filters
            SortFieldsAtoZ = false;	//	Default Sort Order
            UseCustomListsForSorting = true; //	Custom List AutoSort
        }

        public IXLWorksheet Worksheet { get; }

        public IXLPivotTableStyleFormats StyleFormats { get; } = new XLPivotTableStyleFormats();

        public IEnumerable<IXLPivotStyleFormat> AllStyleFormats
        {
            get
            {
                foreach (var styleFormat in this.StyleFormats.RowGrandTotalFormats)
                    yield return styleFormat;

                foreach (var styleFormat in this.StyleFormats.ColumnGrandTotalFormats)
                    yield return styleFormat;

                foreach (var pivotField in ImplementedFields)
                {
                    yield return pivotField.StyleFormats.Subtotal;
                    yield return pivotField.StyleFormats.Header;
                    yield return pivotField.StyleFormats.Label;
                    yield return pivotField.StyleFormats.DataValuesFormat;
                }
            }
        }
    }
}
