using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public enum XLPivotTableTheme
    {
        None,
        PivotStyleDark1,
        PivotStyleDark10,
        PivotStyleDark11,
        PivotStyleDark12,
        PivotStyleDark13,
        PivotStyleDark14,
        PivotStyleDark15,
        PivotStyleDark16,
        PivotStyleDark17,
        PivotStyleDark18,
        PivotStyleDark19,
        PivotStyleDark2,
        PivotStyleDark20,
        PivotStyleDark21,
        PivotStyleDark22,
        PivotStyleDark23,
        PivotStyleDark24,
        PivotStyleDark25,
        PivotStyleDark26,
        PivotStyleDark27,
        PivotStyleDark28,
        PivotStyleDark3,
        PivotStyleDark4,
        PivotStyleDark5,
        PivotStyleDark6,
        PivotStyleDark7,
        PivotStyleDark8,
        PivotStyleDark9,
        PivotStyleLight1,
        PivotStyleLight10,
        PivotStyleLight11,
        PivotStyleLight12,
        PivotStyleLight13,
        PivotStyleLight14,
        PivotStyleLight15,
        PivotStyleLight16,
        PivotStyleLight17,
        PivotStyleLight18,
        PivotStyleLight19,
        PivotStyleLight2,
        PivotStyleLight20,
        PivotStyleLight21,
        PivotStyleLight22,
        PivotStyleLight23,
        PivotStyleLight24,
        PivotStyleLight25,
        PivotStyleLight26,
        PivotStyleLight27,
        PivotStyleLight28,
        PivotStyleLight3,
        PivotStyleLight4,
        PivotStyleLight5,
        PivotStyleLight6,
        PivotStyleLight7,
        PivotStyleLight8,
        PivotStyleLight9,
        PivotStyleMedium1,
        PivotStyleMedium10,
        PivotStyleMedium11,
        PivotStyleMedium12,
        PivotStyleMedium13,
        PivotStyleMedium14,
        PivotStyleMedium15,
        PivotStyleMedium16,
        PivotStyleMedium17,
        PivotStyleMedium18,
        PivotStyleMedium19,
        PivotStyleMedium2,
        PivotStyleMedium20,
        PivotStyleMedium21,
        PivotStyleMedium22,
        PivotStyleMedium23,
        PivotStyleMedium24,
        PivotStyleMedium25,
        PivotStyleMedium26,
        PivotStyleMedium27,
        PivotStyleMedium28,
        PivotStyleMedium3,
        PivotStyleMedium4,
        PivotStyleMedium5,
        PivotStyleMedium6,
        PivotStyleMedium7,
        PivotStyleMedium8,
        PivotStyleMedium9
    }

    public enum XLPivotSortType
    {
        Default = 0,
        Ascending = 1,
        Descending = 2
    }

    public enum XLPivotSubtotals
    {
        DoNotShow,
        AtTop,
        AtBottom
    }

    public enum XLFilterAreaOrder { DownThenOver, OverThenDown }

    public enum XLItemsToRetain { Automatic, None, Max }

    public enum XLPivotTableSourceType { Range, Table }

    public interface IXLPivotTable
    {
        XLPivotTableTheme Theme { get; set; }

        IXLPivotFields Fields { get; }
        IXLPivotFields ReportFilters { get; }
        IXLPivotFields ColumnLabels { get; }
        IXLPivotFields RowLabels { get; }
        IXLPivotValues Values { get; }

        string Name { get; set; }
        string Title { get; set; }
        string Description { get; set; }
        string GrandTotalCaption { get; set; }
        string DataCaption { get; set; }

        string ColumnHeaderCaption { get; set; }
        string RowHeaderCaption { get; set; }

        IXLCell TargetCell { get; set; }

        IXLRange SourceRange { get; set; }
        IXLTable SourceTable { get; set; }
        XLPivotTableSourceType SourceType { get; }

        IEnumerable<string> SourceRangeFieldsAvailable { get; }

        bool MergeAndCenterWithLabels { get; set; } // MergeItem
        int RowLabelIndent { get; set; } // Indent
        XLFilterAreaOrder FilterAreaOrder { get; set; } // PageOverThenDown
        int FilterFieldsPageWrap { get; set; } // PageWrap
        string ErrorValueReplacement { get; set; } // ErrorCaption
        string EmptyCellReplacement { get; set; } // MissingCaption
        bool AutofitColumns { get; set; } //UseAutoFormatting
        bool PreserveCellFormatting { get; set; } // PreserveFormatting

        bool ShowGrandTotalsRows { get; set; } // RowGrandTotals
        bool ShowGrandTotalsColumns { get; set; } // ColumnGrandTotals
        bool FilteredItemsInSubtotals { get; set; } // Subtotal filtered page items
        bool AllowMultipleFilters { get; set; } // MultipleFieldFilters
        bool UseCustomListsForSorting { get; set; } // CustomListSort

        bool ShowExpandCollapseButtons { get; set; }
        bool ShowContextualTooltips { get; set; }
        bool ShowPropertiesInTooltips { get; set; }
        bool DisplayCaptionsAndDropdowns { get; set; }
        bool ClassicPivotTableLayout { get; set; }
        bool ShowValuesRow { get; set; }
        bool ShowEmptyItemsOnRows { get; set; }
        bool ShowEmptyItemsOnColumns { get; set; }
        bool DisplayItemLabels { get; set; }
        bool SortFieldsAtoZ { get; set; }

        bool PrintExpandCollapsedButtons { get; set; }
        bool RepeatRowLabels { get; set; }
        bool PrintTitles { get; set; }

        bool SaveSourceData { get; set; }
        bool EnableShowDetails { get; set; }
        bool RefreshDataOnOpen { get; set; }
        XLItemsToRetain ItemsToRetainPerField { get; set; }
        bool EnableCellEditing { get; set; }

        IXLPivotTable CopyTo(IXLCell targetCell);

        IXLPivotTable SetName(string value);

        IXLPivotTable SetTitle(string value);

        IXLPivotTable SetDescription(string value);
        IXLPivotTable SetMergeAndCenterWithLabels(); IXLPivotTable SetMergeAndCenterWithLabels(bool value);

        IXLPivotTable SetRowLabelIndent(int value);

        IXLPivotTable SetFilterAreaOrder(XLFilterAreaOrder value);

        IXLPivotTable SetFilterFieldsPageWrap(int value);

        IXLPivotTable SetErrorValueReplacement(string value);

        IXLPivotTable SetEmptyCellReplacement(string value);

        IXLPivotTable SetAutofitColumns(); IXLPivotTable SetAutofitColumns(bool value);

        IXLPivotTable SetPreserveCellFormatting(); IXLPivotTable SetPreserveCellFormatting(bool value);

        IXLPivotTable SetShowGrandTotalsRows(); IXLPivotTable SetShowGrandTotalsRows(bool value);

        IXLPivotTable SetShowGrandTotalsColumns(); IXLPivotTable SetShowGrandTotalsColumns(bool value);

        IXLPivotTable SetFilteredItemsInSubtotals(); IXLPivotTable SetFilteredItemsInSubtotals(bool value);

        IXLPivotTable SetAllowMultipleFilters(); IXLPivotTable SetAllowMultipleFilters(bool value);

        IXLPivotTable SetUseCustomListsForSorting(); IXLPivotTable SetUseCustomListsForSorting(bool value);

        IXLPivotTable SetShowExpandCollapseButtons(); IXLPivotTable SetShowExpandCollapseButtons(bool value);

        IXLPivotTable SetShowContextualTooltips(); IXLPivotTable SetShowContextualTooltips(bool value);

        IXLPivotTable SetShowPropertiesInTooltips(); IXLPivotTable SetShowPropertiesInTooltips(bool value);

        IXLPivotTable SetDisplayCaptionsAndDropdowns(); IXLPivotTable SetDisplayCaptionsAndDropdowns(bool value);

        IXLPivotTable SetClassicPivotTableLayout(); IXLPivotTable SetClassicPivotTableLayout(bool value);

        IXLPivotTable SetShowValuesRow(); IXLPivotTable SetShowValuesRow(bool value);

        IXLPivotTable SetShowEmptyItemsOnRows(); IXLPivotTable SetShowEmptyItemsOnRows(bool value);

        IXLPivotTable SetShowEmptyItemsOnColumns(); IXLPivotTable SetShowEmptyItemsOnColumns(bool value);

        IXLPivotTable SetDisplayItemLabels(); IXLPivotTable SetDisplayItemLabels(bool value);

        IXLPivotTable SetSortFieldsAtoZ(); IXLPivotTable SetSortFieldsAtoZ(bool value);

        IXLPivotTable SetPrintExpandCollapsedButtons(); IXLPivotTable SetPrintExpandCollapsedButtons(bool value);

        IXLPivotTable SetRepeatRowLabels(); IXLPivotTable SetRepeatRowLabels(bool value);

        IXLPivotTable SetPrintTitles(); IXLPivotTable SetPrintTitles(bool value);

        IXLPivotTable SetSaveSourceData(); IXLPivotTable SetSaveSourceData(bool value);

        IXLPivotTable SetEnableShowDetails(); IXLPivotTable SetEnableShowDetails(bool value);

        IXLPivotTable SetRefreshDataOnOpen(); IXLPivotTable SetRefreshDataOnOpen(bool value);

        IXLPivotTable SetItemsToRetainPerField(XLItemsToRetain value);

        IXLPivotTable SetEnableCellEditing(); IXLPivotTable SetEnableCellEditing(bool value);

        IXLPivotTable SetColumnHeaderCaption(string value);

        IXLPivotTable SetRowHeaderCaption(string value);

        bool ShowRowHeaders { get; set; }
        bool ShowColumnHeaders { get; set; }
        bool ShowRowStripes { get; set; }
        bool ShowColumnStripes { get; set; }
        XLPivotSubtotals Subtotals { get; set; }
        XLPivotLayout Layout { set; }
        bool InsertBlankLines { set; }

        IXLPivotTable SetShowRowHeaders(); IXLPivotTable SetShowRowHeaders(bool value);

        IXLPivotTable SetShowColumnHeaders(); IXLPivotTable SetShowColumnHeaders(bool value);

        IXLPivotTable SetShowRowStripes(); IXLPivotTable SetShowRowStripes(bool value);

        IXLPivotTable SetShowColumnStripes(); IXLPivotTable SetShowColumnStripes(bool value);

        IXLPivotTable SetSubtotals(XLPivotSubtotals value);

        IXLPivotTable SetLayout(XLPivotLayout value);

        IXLPivotTable SetInsertBlankLines(); IXLPivotTable SetInsertBlankLines(bool value);

        IXLWorksheet Worksheet { get; }

        IXLPivotTableStyleFormats StyleFormats { get; }
    }
}
