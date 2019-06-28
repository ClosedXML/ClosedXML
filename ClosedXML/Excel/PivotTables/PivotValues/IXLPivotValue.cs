// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    public enum XLPivotCalculation
    {
        Normal,
        DifferenceFrom,
        PercentageOf,
        PercentageDifferenceFrom,
        RunningTotal,
        PercentageOfRow,
        PercentageOfColumn,
        PercentageOfTotal,
        Index
    }

    public enum XLPivotCalculationItem
    {
        Value, Previous, Next
    }

    public enum XLPivotSummary
    {
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

    public interface IXLPivotValue
    {
        /// <summary>
        /// Specifies the index to the base field when the ShowDataAs calculation is in use.
        /// </summary>
        /// <value>
        /// The name of the column of the relevant base field.
        /// </value>
        String BaseFieldName { get; set; }

        /// <summary>
        /// Specifies the index to the base item when the ShowDataAs calculation is in use.
        /// </summary>
        /// <value>
        /// The value of the referenced base field item.
        /// </value>
        Object BaseItemValue { get; set; }

        XLPivotCalculation Calculation { get; set; }
        XLPivotCalculationItem CalculationItem { get; set; }
        String CustomName { get; set; }
        IXLPivotValueFormat NumberFormat { get; }
        String SourceName { get; }
        XLPivotSummary SummaryFormula { get; set; }

        IXLPivotValue SetBaseFieldName(String value);

        IXLPivotValue SetBaseItemValue(Object value);

        IXLPivotValue SetCalculation(XLPivotCalculation value);

        IXLPivotValue SetCalculationItem(XLPivotCalculationItem value);

        IXLPivotValue SetSummaryFormula(XLPivotSummary value);

        IXLPivotValueCombination ShowAsDifferenceFrom(String fieldSourceName);

        IXLPivotValue ShowAsIndex();

        IXLPivotValue ShowAsNormal();

        IXLPivotValueCombination ShowAsPercentageDifferenceFrom(String fieldSourceName);

        IXLPivotValueCombination ShowAsPercentageFrom(String fieldSourceName);

        IXLPivotValue ShowAsPercentageOfColumn();

        IXLPivotValue ShowAsPercentageOfRow();

        IXLPivotValue ShowAsPercentageOfTotal();

        IXLPivotValue ShowAsRunningTotalIn(String fieldSourceName);
    }
}
