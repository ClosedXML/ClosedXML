using System;
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

        String Name { get; set; }
        String Title { get; set; }
        String Description { get; set; }

        String ColumnHeaderCaption { get; set; }
        String RowHeaderCaption { get; set; }

        IXLCell TargetCell { get; set; }

        IXLRange SourceRange { get; set; }
        IXLTable SourceTable { get; set; }
        XLPivotTableSourceType SourceType { get; }

        IEnumerable<String> SourceRangeFieldsAvailable { get; }

        Boolean MergeAndCenterWithLabels { get; set; } // MergeItem
        Int32 RowLabelIndent { get; set; } // Indent
        XLFilterAreaOrder FilterAreaOrder { get; set; } // PageOverThenDown
        Int32 FilterFieldsPageWrap { get; set; } // PageWrap
        String ErrorValueReplacement { get; set; } // ErrorCaption
        String EmptyCellReplacement { get; set; } // MissingCaption
        Boolean AutofitColumns { get; set; } //UseAutoFormatting
        Boolean PreserveCellFormatting { get; set; } // PreserveFormatting

        Boolean ShowGrandTotalsRows { get; set; } // RowGrandTotals
        Boolean ShowGrandTotalsColumns { get; set; } // ColumnGrandTotals
        Boolean FilteredItemsInSubtotals { get; set; } // Subtotal filtered page items
        Boolean AllowMultipleFilters { get; set; } // MultipleFieldFilters
        Boolean UseCustomListsForSorting { get; set; } // CustomListSort

        Boolean ShowExpandCollapseButtons { get; set; }
        Boolean ShowContextualTooltips { get; set; }
        Boolean ShowPropertiesInTooltips { get; set; }
        Boolean DisplayCaptionsAndDropdowns { get; set; }
        Boolean ClassicPivotTableLayout { get; set; }
        Boolean ShowValuesRow { get; set; }
        Boolean ShowEmptyItemsOnRows { get; set; }
        Boolean ShowEmptyItemsOnColumns { get; set; }
        Boolean DisplayItemLabels { get; set; }
        Boolean SortFieldsAtoZ { get; set; }

        Boolean PrintExpandCollapsedButtons { get; set; }
        Boolean RepeatRowLabels { get; set; }
        Boolean PrintTitles { get; set; }

        Boolean SaveSourceData { get; set; }
        Boolean EnableShowDetails { get; set; }
        Boolean RefreshDataOnOpen { get; set; }
        XLItemsToRetain ItemsToRetainPerField { get; set; }
        Boolean EnableCellEditing { get; set; }

        IXLPivotTable CopyTo(IXLCell targetCell);

        IXLPivotTable SetName(String value);

        IXLPivotTable SetTitle(String value);

        IXLPivotTable SetDescription(String value);

        IXLPivotTable SetMergeAndCenterWithLabels(); IXLPivotTable SetMergeAndCenterWithLabels(Boolean value);

        IXLPivotTable SetRowLabelIndent(Int32 value);

        IXLPivotTable SetFilterAreaOrder(XLFilterAreaOrder value);

        IXLPivotTable SetFilterFieldsPageWrap(Int32 value);

        IXLPivotTable SetErrorValueReplacement(String value);

        IXLPivotTable SetEmptyCellReplacement(String value);

        IXLPivotTable SetAutofitColumns(); IXLPivotTable SetAutofitColumns(Boolean value);

        IXLPivotTable SetPreserveCellFormatting(); IXLPivotTable SetPreserveCellFormatting(Boolean value);

        IXLPivotTable SetShowGrandTotalsRows(); IXLPivotTable SetShowGrandTotalsRows(Boolean value);

        IXLPivotTable SetShowGrandTotalsColumns(); IXLPivotTable SetShowGrandTotalsColumns(Boolean value);

        IXLPivotTable SetFilteredItemsInSubtotals(); IXLPivotTable SetFilteredItemsInSubtotals(Boolean value);

        IXLPivotTable SetAllowMultipleFilters(); IXLPivotTable SetAllowMultipleFilters(Boolean value);

        IXLPivotTable SetUseCustomListsForSorting(); IXLPivotTable SetUseCustomListsForSorting(Boolean value);

        IXLPivotTable SetShowExpandCollapseButtons(); IXLPivotTable SetShowExpandCollapseButtons(Boolean value);

        IXLPivotTable SetShowContextualTooltips(); IXLPivotTable SetShowContextualTooltips(Boolean value);

        IXLPivotTable SetShowPropertiesInTooltips(); IXLPivotTable SetShowPropertiesInTooltips(Boolean value);

        IXLPivotTable SetDisplayCaptionsAndDropdowns(); IXLPivotTable SetDisplayCaptionsAndDropdowns(Boolean value);

        IXLPivotTable SetClassicPivotTableLayout(); IXLPivotTable SetClassicPivotTableLayout(Boolean value);

        IXLPivotTable SetShowValuesRow(); IXLPivotTable SetShowValuesRow(Boolean value);

        IXLPivotTable SetShowEmptyItemsOnRows(); IXLPivotTable SetShowEmptyItemsOnRows(Boolean value);

        IXLPivotTable SetShowEmptyItemsOnColumns(); IXLPivotTable SetShowEmptyItemsOnColumns(Boolean value);

        IXLPivotTable SetDisplayItemLabels(); IXLPivotTable SetDisplayItemLabels(Boolean value);

        IXLPivotTable SetSortFieldsAtoZ(); IXLPivotTable SetSortFieldsAtoZ(Boolean value);

        IXLPivotTable SetPrintExpandCollapsedButtons(); IXLPivotTable SetPrintExpandCollapsedButtons(Boolean value);

        IXLPivotTable SetRepeatRowLabels(); IXLPivotTable SetRepeatRowLabels(Boolean value);

        IXLPivotTable SetPrintTitles(); IXLPivotTable SetPrintTitles(Boolean value);

        IXLPivotTable SetSaveSourceData(); IXLPivotTable SetSaveSourceData(Boolean value);

        IXLPivotTable SetEnableShowDetails(); IXLPivotTable SetEnableShowDetails(Boolean value);

        IXLPivotTable SetRefreshDataOnOpen(); IXLPivotTable SetRefreshDataOnOpen(Boolean value);

        IXLPivotTable SetItemsToRetainPerField(XLItemsToRetain value);

        IXLPivotTable SetEnableCellEditing(); IXLPivotTable SetEnableCellEditing(Boolean value);

        IXLPivotTable SetColumnHeaderCaption(String value);

        IXLPivotTable SetRowHeaderCaption(String value);

        Boolean ShowRowHeaders { get; set; }
        Boolean ShowColumnHeaders { get; set; }
        Boolean ShowRowStripes { get; set; }
        Boolean ShowColumnStripes { get; set; }
        XLPivotSubtotals Subtotals { get; set; }
        XLPivotLayout Layout { set; }
        Boolean InsertBlankLines { set; }

        IXLPivotTable SetShowRowHeaders(); IXLPivotTable SetShowRowHeaders(Boolean value);

        IXLPivotTable SetShowColumnHeaders(); IXLPivotTable SetShowColumnHeaders(Boolean value);

        IXLPivotTable SetShowRowStripes(); IXLPivotTable SetShowRowStripes(Boolean value);

        IXLPivotTable SetShowColumnStripes(); IXLPivotTable SetShowColumnStripes(Boolean value);

        IXLPivotTable SetSubtotals(XLPivotSubtotals value);

        IXLPivotTable SetLayout(XLPivotLayout value);

        IXLPivotTable SetInsertBlankLines(); IXLPivotTable SetInsertBlankLines(Boolean value);

        IXLWorksheet Worksheet { get; }

        IXLPivotTableStyleFormats StyleFormats { get; }
    }
}
