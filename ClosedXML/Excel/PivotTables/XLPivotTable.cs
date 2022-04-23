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
        private string _name;
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
            newPivotTable.DataCaption = DataCaption;
            newPivotTable.GrandTotalCaption = GrandTotalCaption;
            // TODO: Copy Styleformats

            return newPivotTable;
        }

        public IXLPivotTable SetTheme(XLPivotTableTheme value)
        {
            Theme = value; return this;
        }

        public string Name
        {
            get { return _name; }
            set
            {
                if (_name == value) return;

                var oldname = _name ?? string.Empty;

                if (!XLHelper.ValidateName("pivot table", value, oldname, Worksheet.PivotTables.Select(pvt => pvt.Name), out string message))
                    throw new ArgumentException(message, nameof(value));

                _name = value;

                if (!string.IsNullOrWhiteSpace(oldname) && !string.Equals(oldname, _name, StringComparison.OrdinalIgnoreCase))
                {
                    Worksheet.PivotTables.Delete(oldname);
                    (Worksheet.PivotTables as XLPivotTables).Add(_name, this);
                }
            }
        }

        public IXLPivotTable SetName(string value)
        {
            Name = value; return this;
        }

        public string Title { get; set; }

        public IXLPivotTable SetTitle(string value)
        {
            Title = value; return this;
        }

        public string Description { get; set; }

        public IXLPivotTable SetDescription(string value)
        {
            Description = value; return this;
        }
        
        public string GrandTotalCaption { get; set; }

        public string DataCaption { get; set; }

        public string ColumnHeaderCaption { get; set; }

        public IXLPivotTable SetColumnHeaderCaption(string value)
        {
            ColumnHeaderCaption = value;
            return this;
        }

        public string RowHeaderCaption { get; set; }

        public IXLPivotTable SetRowHeaderCaption(string value)
        {
            RowHeaderCaption = value;
            return this;
        }

        public bool MergeAndCenterWithLabels { get; set; }

        public IXLPivotTable SetMergeAndCenterWithLabels()
        {
            MergeAndCenterWithLabels = true; return this;
        }

        public IXLPivotTable SetMergeAndCenterWithLabels(bool value)
        {
            MergeAndCenterWithLabels = value; return this;
        }

        public int RowLabelIndent { get; set; }

        public IXLPivotTable SetRowLabelIndent(int value)
        {
            RowLabelIndent = value; return this;
        }

        public XLFilterAreaOrder FilterAreaOrder { get; set; }

        public IXLPivotTable SetFilterAreaOrder(XLFilterAreaOrder value)
        {
            FilterAreaOrder = value; return this;
        }

        public int FilterFieldsPageWrap { get; set; }

        public IXLPivotTable SetFilterFieldsPageWrap(int value)
        {
            FilterFieldsPageWrap = value; return this;
        }

        public string ErrorValueReplacement { get; set; }

        public IXLPivotTable SetErrorValueReplacement(string value)
        {
            ErrorValueReplacement = value; return this;
        }

        public string EmptyCellReplacement { get; set; }

        public IXLPivotTable SetEmptyCellReplacement(string value)
        {
            EmptyCellReplacement = value; return this;
        }

        public bool AutofitColumns { get; set; }

        public IXLPivotTable SetAutofitColumns()
        {
            AutofitColumns = true; return this;
        }

        public IXLPivotTable SetAutofitColumns(bool value)
        {
            AutofitColumns = value; return this;
        }

        public bool PreserveCellFormatting { get; set; }

        public IXLPivotTable SetPreserveCellFormatting()
        {
            PreserveCellFormatting = true; return this;
        }

        public IXLPivotTable SetPreserveCellFormatting(bool value)
        {
            PreserveCellFormatting = value; return this;
        }

        public bool ShowGrandTotalsRows { get; set; }

        public IXLPivotTable SetShowGrandTotalsRows()
        {
            ShowGrandTotalsRows = true; return this;
        }

        public IXLPivotTable SetShowGrandTotalsRows(bool value)
        {
            ShowGrandTotalsRows = value; return this;
        }

        public bool ShowGrandTotalsColumns { get; set; }

        public IXLPivotTable SetShowGrandTotalsColumns()
        {
            ShowGrandTotalsColumns = true; return this;
        }

        public IXLPivotTable SetShowGrandTotalsColumns(bool value)
        {
            ShowGrandTotalsColumns = value; return this;
        }

        public bool FilteredItemsInSubtotals { get; set; }

        public IXLPivotTable SetFilteredItemsInSubtotals()
        {
            FilteredItemsInSubtotals = true; return this;
        }

        public IXLPivotTable SetFilteredItemsInSubtotals(bool value)
        {
            FilteredItemsInSubtotals = value; return this;
        }

        public bool AllowMultipleFilters { get; set; }

        public IXLPivotTable SetAllowMultipleFilters()
        {
            AllowMultipleFilters = true; return this;
        }

        public IXLPivotTable SetAllowMultipleFilters(bool value)
        {
            AllowMultipleFilters = value; return this;
        }

        public bool UseCustomListsForSorting { get; set; }

        public IXLPivotTable SetUseCustomListsForSorting()
        {
            UseCustomListsForSorting = true; return this;
        }

        public IXLPivotTable SetUseCustomListsForSorting(bool value)
        {
            UseCustomListsForSorting = value; return this;
        }

        public bool ShowExpandCollapseButtons { get; set; }

        public IXLPivotTable SetShowExpandCollapseButtons()
        {
            ShowExpandCollapseButtons = true; return this;
        }

        public IXLPivotTable SetShowExpandCollapseButtons(bool value)
        {
            ShowExpandCollapseButtons = value; return this;
        }

        public bool ShowContextualTooltips { get; set; }

        public IXLPivotTable SetShowContextualTooltips()
        {
            ShowContextualTooltips = true; return this;
        }

        public IXLPivotTable SetShowContextualTooltips(bool value)
        {
            ShowContextualTooltips = value; return this;
        }

        public bool ShowPropertiesInTooltips { get; set; }

        public IXLPivotTable SetShowPropertiesInTooltips()
        {
            ShowPropertiesInTooltips = true; return this;
        }

        public IXLPivotTable SetShowPropertiesInTooltips(bool value)
        {
            ShowPropertiesInTooltips = value; return this;
        }

        public bool DisplayCaptionsAndDropdowns { get; set; }

        public IXLPivotTable SetDisplayCaptionsAndDropdowns()
        {
            DisplayCaptionsAndDropdowns = true; return this;
        }

        public IXLPivotTable SetDisplayCaptionsAndDropdowns(bool value)
        {
            DisplayCaptionsAndDropdowns = value; return this;
        }

        public bool ClassicPivotTableLayout { get; set; }

        public IXLPivotTable SetClassicPivotTableLayout()
        {
            ClassicPivotTableLayout = true; return this;
        }

        public IXLPivotTable SetClassicPivotTableLayout(bool value)
        {
            ClassicPivotTableLayout = value; return this;
        }

        public bool ShowValuesRow { get; set; }

        public IXLPivotTable SetShowValuesRow()
        {
            ShowValuesRow = true; return this;
        }

        public IXLPivotTable SetShowValuesRow(bool value)
        {
            ShowValuesRow = value; return this;
        }

        public bool ShowEmptyItemsOnRows { get; set; }

        public IXLPivotTable SetShowEmptyItemsOnRows()
        {
            ShowEmptyItemsOnRows = true; return this;
        }

        public IXLPivotTable SetShowEmptyItemsOnRows(bool value)
        {
            ShowEmptyItemsOnRows = value; return this;
        }

        public bool ShowEmptyItemsOnColumns { get; set; }

        public IXLPivotTable SetShowEmptyItemsOnColumns()
        {
            ShowEmptyItemsOnColumns = true; return this;
        }

        public IXLPivotTable SetShowEmptyItemsOnColumns(bool value)
        {
            ShowEmptyItemsOnColumns = value; return this;
        }

        public bool DisplayItemLabels { get; set; }

        public IXLPivotTable SetDisplayItemLabels()
        {
            DisplayItemLabels = true; return this;
        }

        public IXLPivotTable SetDisplayItemLabels(bool value)
        {
            DisplayItemLabels = value; return this;
        }

        public bool SortFieldsAtoZ { get; set; }

        public IXLPivotTable SetSortFieldsAtoZ()
        {
            SortFieldsAtoZ = true; return this;
        }

        public IXLPivotTable SetSortFieldsAtoZ(bool value)
        {
            SortFieldsAtoZ = value; return this;
        }

        public bool PrintExpandCollapsedButtons { get; set; }

        public IXLPivotTable SetPrintExpandCollapsedButtons()
        {
            PrintExpandCollapsedButtons = true; return this;
        }

        public IXLPivotTable SetPrintExpandCollapsedButtons(bool value)
        {
            PrintExpandCollapsedButtons = value; return this;
        }

        public bool RepeatRowLabels { get; set; }

        public IXLPivotTable SetRepeatRowLabels()
        {
            RepeatRowLabels = true; return this;
        }

        public IXLPivotTable SetRepeatRowLabels(bool value)
        {
            RepeatRowLabels = value; return this;
        }

        public bool PrintTitles { get; set; }

        public IXLPivotTable SetPrintTitles()
        {
            PrintTitles = true; return this;
        }

        public IXLPivotTable SetPrintTitles(bool value)
        {
            PrintTitles = value; return this;
        }

        public bool SaveSourceData { get; set; }

        public IXLPivotTable SetSaveSourceData()
        {
            SaveSourceData = true; return this;
        }

        public IXLPivotTable SetSaveSourceData(bool value)
        {
            SaveSourceData = value; return this;
        }

        public bool EnableShowDetails { get; set; }

        public IXLPivotTable SetEnableShowDetails()
        {
            EnableShowDetails = true; return this;
        }

        public IXLPivotTable SetEnableShowDetails(bool value)
        {
            EnableShowDetails = value; return this;
        }

        public bool RefreshDataOnOpen { get; set; }

        public IXLPivotTable SetRefreshDataOnOpen()
        {
            RefreshDataOnOpen = true; return this;
        }

        public IXLPivotTable SetRefreshDataOnOpen(bool value)
        {
            RefreshDataOnOpen = value; return this;
        }

        public XLItemsToRetain ItemsToRetainPerField { get; set; }

        public IXLPivotTable SetItemsToRetainPerField(XLItemsToRetain value)
        {
            ItemsToRetainPerField = value; return this;
        }

        public bool EnableCellEditing { get; set; }

        public IXLPivotTable SetEnableCellEditing()
        {
            EnableCellEditing = true; return this;
        }

        public IXLPivotTable SetEnableCellEditing(bool value)
        {
            EnableCellEditing = value; return this;
        }

        public bool ShowRowHeaders { get; set; }

        public IXLPivotTable SetShowRowHeaders()
        {
            ShowRowHeaders = true; return this;
        }

        public IXLPivotTable SetShowRowHeaders(bool value)
        {
            ShowRowHeaders = value; return this;
        }

        public bool ShowColumnHeaders { get; set; }

        public IXLPivotTable SetShowColumnHeaders()
        {
            ShowColumnHeaders = true; return this;
        }

        public IXLPivotTable SetShowColumnHeaders(bool value)
        {
            ShowColumnHeaders = value; return this;
        }

        public bool ShowRowStripes { get; set; }

        public IXLPivotTable SetShowRowStripes()
        {
            ShowRowStripes = true; return this;
        }

        public IXLPivotTable SetShowRowStripes(bool value)
        {
            ShowRowStripes = value; return this;
        }

        public bool ShowColumnStripes { get; set; }

        public IXLPivotTable SetShowColumnStripes()
        {
            ShowColumnStripes = true; return this;
        }

        public IXLPivotTable SetShowColumnStripes(bool value)
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

        public bool InsertBlankLines
        {
            set { Fields.ForEach(f => f.SetInsertBlankLines(value)); }
        }

        public IXLPivotTable SetInsertBlankLines()
        {
            InsertBlankLines = true; return this;
        }

        public IXLPivotTable SetInsertBlankLines(bool value)
        {
            InsertBlankLines = value; return this;
        }

        internal string RelId { get; set; }
        internal string CacheDefinitionRelId { get; set; }
        internal string WorkbookCacheRelId { get; set; }

        private void SetExcelDefaults()
        {
            EmptyCellReplacement = string.Empty;
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
