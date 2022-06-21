using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public enum XLSubtotalFunction
    {
        /// <summary>
        /// A subtotal function to display a sum for a pivot field. The value is displayed only for a field if
        /// no other subtotal is specified.
        /// </summary>
        Automatic,
        None,
        Sum,
        Count,
        Average,
        Minimum,
        Maximum,
        Product,
        CountNumbers,
        StandardDeviation,
        PopulationStandardDeviation,
        Variance,
        PopulationVariance,
    }

    public enum XLPivotLayout { Outline, Tabular, Compact }

    public interface IXLPivotField
    {
        String SourceName { get; }
        String CustomName { get; set; }
        String SubtotalCaption { get; set; }

        /// <summary>
        /// Subtotal functions that should be displayed for the field. If empty, no subtotal is displayed for the field.
        /// It's not possible to modify the order of displayed subtotal rows in the pivot table.
        /// </summary>
        IEnumerable<XLSubtotalFunction> Subtotals { get; }

        Boolean IncludeNewItemsInFilter { get; set; }

        Boolean Outline { get; set; }
        Boolean Compact { get; set; }
        Boolean? SubtotalsAtTop { get; set; }
        Boolean RepeatItemLabels { get; set; }
        Boolean InsertBlankLines { get; set; }
        Boolean ShowBlankItems { get; set; }
        Boolean InsertPageBreaks { get; set; }
        Boolean Collapsed { get; set; }
        XLPivotSortType SortType { get; set; }

        IXLPivotField SetCustomName(String value);

        IXLPivotField SetSubtotalCaption(String value);

        /// <summary>
        /// Change if a subtotal function should be displayed for the pivot field. If the pivot field already contains same subtotal function
        /// it won't be added for second time.
        /// Function <see cref="XLSubtotalFunction.Automatic"/> is a fallback and won't be displayed, if any other subtotal is set.
        /// </summary>
        /// <param name="function">A subtotal function to change.</param>
        /// <param name="enabled">Should the subtotal function be included in subtotals of the field or not.</param>
        IXLPivotField SetSubtotal(XLSubtotalFunction function, bool enabled);

        IXLPivotField SetIncludeNewItemsInFilter(); IXLPivotField SetIncludeNewItemsInFilter(Boolean value);

        IXLPivotField SetLayout(XLPivotLayout value);

        IXLPivotField SetSubtotalsAtTop(); IXLPivotField SetSubtotalsAtTop(Boolean value);

        IXLPivotField SetRepeatItemLabels(); IXLPivotField SetRepeatItemLabels(Boolean value);

        IXLPivotField SetInsertBlankLines(); IXLPivotField SetInsertBlankLines(Boolean value);

        IXLPivotField SetShowBlankItems(); IXLPivotField SetShowBlankItems(Boolean value);

        IXLPivotField SetInsertPageBreaks(); IXLPivotField SetInsertPageBreaks(Boolean value);

        IXLPivotField SetCollapsed(); IXLPivotField SetCollapsed(Boolean value);

        IXLPivotField SetSort(XLPivotSortType value);

        IList<Object> SelectedValues { get; }

        IXLPivotField AddSelectedValue(Object value);

        IXLPivotField AddSelectedValues(IEnumerable<Object> values);

        IXLPivotFieldStyleFormats StyleFormats { get; }

        Boolean IsOnRowAxis { get; }
        Boolean IsOnColumnAxis { get; }
        Boolean IsInFilterList { get; }

        Int32 Offset { get; }
    }
}
