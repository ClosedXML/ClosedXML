using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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
    public enum XLLabelForm { Outline, Tabular }
    public interface IXLPivotField
    {
        String SourceName { get; }
        String CustomName { get; set; }

        XLSubtotalFunction Subtotals { get; set; }
        Boolean IncludeNewItemsInFilter { get; set; }

        XLLabelForm ItemLabelsForm { get; set; }
        Boolean CompactForm { get; set; }
        Boolean SubtotalsAtTop { get; set; }
        Boolean RepeatItemLabels { get; set; }
        Boolean InsertBlankLines  { get; set; }
        Boolean ShowBlankItems { get; set; }
        Boolean InsertPageBreaks { get; set; }

        IXLPivotField SetCustomName(String value);

        IXLPivotField SetSubtotals(XLSubtotalFunction value);
        IXLPivotField SetIncludeNewItemsInFilter(); IXLPivotField SetIncludeNewItemsInFilter(Boolean value);

        IXLPivotField SetItemLabelsForm(XLLabelForm value);
        IXLPivotField SetCompactForm(); IXLPivotField SetCompactForm(Boolean value);
        IXLPivotField SetSubtotalsAtTop(); IXLPivotField SetSubtotalsAtTop(Boolean value);
        IXLPivotField SetRepeatItemLabels(); IXLPivotField SetRepeatItemLabels(Boolean value);
        IXLPivotField SetInsertBlankLines(); IXLPivotField SetInsertBlankLines(Boolean value);
        IXLPivotField SetShowBlankItems(); IXLPivotField SetShowBlankItems(Boolean value);
        IXLPivotField SetInsertPageBreaks(); IXLPivotField SetInsertPageBreaks(Boolean value);

    }
}
