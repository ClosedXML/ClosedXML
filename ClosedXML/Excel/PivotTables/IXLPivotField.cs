using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public enum XLSubtotalFunction
    {
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
        string SourceName { get; }
        string CustomName { get; set; }
        string SubtotalCaption { get; set; }

        List<XLSubtotalFunction> Subtotals { get; }
        bool IncludeNewItemsInFilter { get; set; }

        bool Outline { get; set; }
        bool Compact { get; set; }
        bool? SubtotalsAtTop { get; set; }
        bool RepeatItemLabels { get; set; }
        bool InsertBlankLines { get; set; }
        bool ShowBlankItems { get; set; }
        bool InsertPageBreaks { get; set; }
        bool Collapsed { get; set; }
        XLPivotSortType SortType { get; set; }

        IXLPivotField SetCustomName(string value);

        IXLPivotField SetSubtotalCaption(string value);

        IXLPivotField AddSubtotal(XLSubtotalFunction value);

        IXLPivotField SetIncludeNewItemsInFilter(); IXLPivotField SetIncludeNewItemsInFilter(bool value);

        IXLPivotField SetLayout(XLPivotLayout value);

        IXLPivotField SetSubtotalsAtTop(); IXLPivotField SetSubtotalsAtTop(bool value);

        IXLPivotField SetRepeatItemLabels(); IXLPivotField SetRepeatItemLabels(bool value);

        IXLPivotField SetInsertBlankLines(); IXLPivotField SetInsertBlankLines(bool value);

        IXLPivotField SetShowBlankItems(); IXLPivotField SetShowBlankItems(bool value);

        IXLPivotField SetInsertPageBreaks(); IXLPivotField SetInsertPageBreaks(bool value);

        IXLPivotField SetCollapsed(); IXLPivotField SetCollapsed(bool value);

        IXLPivotField SetSort(XLPivotSortType value);

        IList<object> SelectedValues { get; }

        IXLPivotField AddSelectedValue(object value);

        IXLPivotField AddSelectedValues(IEnumerable<object> values);

        IXLPivotFieldStyleFormats StyleFormats { get; }

        bool IsOnRowAxis { get; }
        bool IsOnColumnAxis { get; }
        bool IsInFilterList { get; }

        int Offset { get; }
    }
}
