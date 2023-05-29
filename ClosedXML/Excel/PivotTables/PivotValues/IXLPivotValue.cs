#nullable disable

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
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

    /// <summary>
    /// A pivot value field, it is basically a specification of how to determine and
    /// format values from source to display in the pivot table.
    /// </summary>
    public interface IXLPivotValue
    {
        String SourceName { get; }
        String CustomName { get; set; }

        IXLPivotValueFormat NumberFormat { get; }

        XLPivotSummary SummaryFormula { get; set; }
        XLPivotCalculation Calculation { get; set; }

        /// <summary>
        /// Name of a base field to calculate a value to show in the pivot table. The base field determines which
        /// base items can be used. Instead of base item, previous or next value can be used through <see cref="CalculationItem" />
        /// </summary>
        /// <remarks>Used only if the value should be showed <b>Show Values As</b> in the value field settings.</remarks>
        /// <example>
        /// Show values as a percent of a specific value of a different field, e.g. as a % of units sold from Q1 (quarts is a base field and Q1 is a base item).
        /// </example>
        String BaseField { get; set; }

        /// <summary>
        /// The value of a base item to calculate a value to show in the pivot table. The base item is selected from values of a base field.
        /// </summary>
        /// <remarks>Used only if the value should be showed <b>Show Values As</b> in the value field settings.</remarks>
        /// <example>
        /// Show values as a percent of a specific value of a different field, e.g. as a % of units sold from Q1 (quarts is a base field and Q1 is a base item).
        /// </example>
        XLCellValue BaseItem { get; set; }

        XLPivotCalculationItem CalculationItem { get; set; }

        IXLPivotValue SetSummaryFormula(XLPivotSummary value);
        IXLPivotValue SetCalculation(XLPivotCalculation value);
        IXLPivotValue SetBaseField(String value);
        IXLPivotValue SetBaseItem(XLCellValue value);
        IXLPivotValue SetCalculationItem(XLPivotCalculationItem value);


        IXLPivotValue ShowAsNormal();
        IXLPivotValueCombination ShowAsDifferenceFrom(String fieldSourceName);
        IXLPivotValueCombination ShowAsPercentageFrom(String fieldSourceName);
        IXLPivotValueCombination ShowAsPercentageDifferenceFrom(String fieldSourceName);
        IXLPivotValue ShowAsRunningTotalIn(String fieldSourceName);
        IXLPivotValue ShowAsPercentageOfRow();
        IXLPivotValue ShowAsPercentageOfColumn();
        IXLPivotValue ShowAsPercentageOfTotal();
        IXLPivotValue ShowAsIndex();

    }
}
