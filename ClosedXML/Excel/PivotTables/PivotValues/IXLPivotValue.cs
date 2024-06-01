#nullable disable

using System;

namespace ClosedXML.Excel
{
    /// <summary>
    /// Enum describing how is a pivot field values (i.e. in data area) displayed.
    /// </summary>
    /// <remarks>
    /// [ISO-29500] 18.18.70 ST_ShowDataAs
    /// </remarks>
    public enum XLPivotCalculation
    {
        /// <summary>
        /// Field values are displayed normally.
        /// </summary>
        Normal,
        DifferenceFrom,
        PercentageOf,
        PercentageDifferenceFrom,
        RunningTotal,
        PercentageOfRow,
        PercentageOfColumn,
        PercentageOfTotal,

        /// <summary>
        /// Basically a relative importance of a value. Closer the value to 1.0 is, the less
        /// important it is. Calculated as <c>(value-in-cell * grand-total-of-grand-totals) /
        /// (grand-total-row * grand-total-column)</c>.
        /// </summary>
        Index
    }

    /// <summary>
    /// Some calculation from <see cref="XLPivotCalculation"/> need a value as another an argument
    /// of a calculation (e.g. difference from). This enum specifies how to find the reference value.
    /// </summary>
    public enum XLPivotCalculationItem
    {
        Value, Previous, Next
    }

    /// <summary>
    /// An enum that specifies how are grouped pivot field values summed up in a single cell of a
    /// pivot table.
    /// </summary>
    /// <remarks>
    /// [ISO-29500] 18.18.17 ST_DataConsolidateFunction
    /// </remarks>
    public enum XLPivotSummary
    {
        /// <summary>
        /// Values are summed up.
        /// </summary>
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

        /// <summary>
        /// Get custom name of pivot value. If custom name is not specified, return source name as
        /// a fallback.
        /// </summary>
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
