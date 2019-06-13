#nullable disable

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

    /// <summary>
    /// A pivot value field, it is basically a specification of how to determine and
    /// format values from source to display in the pivot table.
    /// </summary>
    public interface IXLPivotValue
    {
        /// <summary>
        /// Specifies the index to the base field when the ShowDataAs calculation is in use.
        /// Instead of base item, previous or next value can be used through <see cref="CalculationItem" />
        /// </summary>
        /// <remarks>Used only if the value should be showed <b>Show Values As</b> in the value field settings.</remarks>
        /// <value>
        /// The name of the column of the relevant base field.
        /// </value>
        /// <example>
        /// Show values as a percent of a specific value of a different field, e.g. as a % of units sold from Q1 (quarts is a base field and Q1 is a base item).
        /// </example>
        String BaseFieldName { get; set; }

        /// <summary>
        /// The value of a base item to calculate a value to show in the pivot table. The base item is selected from values of a base field.
        /// </summary>
        /// <remarks>Used only if the value should be showed <b>Show Values As</b> in the value field settings.</remarks>
        /// <value>
        /// The value of the referenced base field item.
        /// </value>
        /// <example>
        /// Show values as a percent of a specific value of a different field, e.g. as a % of units sold from Q1 (quarts is a base field and Q1 is a base item).
        /// </example>
        XLCellValue BaseItemValue { get; set; }

        XLPivotCalculation Calculation { get; set; }
        XLPivotCalculationItem CalculationItem { get; set; }
        String CustomName { get; set; }
        IXLPivotValueFormat NumberFormat { get; }
        String SourceName { get; }
        XLPivotSummary SummaryFormula { get; set; }

        IXLPivotValue SetBaseFieldName(String value);

        IXLPivotValue SetBaseItemValue(XLCellValue value);

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
