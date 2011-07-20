using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public enum XLFilterAreaOrder { DownThenOver, OverThenDown }
    public enum XLItemsToRetain { Automatic, None, Max }

    public interface IXLPivotTable
    {
        IXLPivotFields ReportFilters { get; }
        IXLPivotFields  ColumnLabels { get; }
        IXLPivotFields RowLabels { get; }
        IXLPivotValues Values { get; }

        String Name { get; set; }
        String Title { get; set; }
        String Description { get; set; }

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



    }
}
