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

    public interface IXLPivotValue
    {
        String SourceName { get; }
        String CustomName { get; set; }

        IXLPivotValueFormat NumberFormat { get; }

        XLPivotSummary SummaryFormula { get; set; }
        XLPivotCalculation Calculation { get; set; }
        String BaseField { get; set; }
        String BaseItem { get; set; }
        XLPivotCalculationItem CalculationItem { get; set; }

        IXLPivotValue SetSummaryFormula(XLPivotSummary value);
        IXLPivotValue SetCalculation(XLPivotCalculation value);
        IXLPivotValue SetBaseField(String value);
        IXLPivotValue SetBaseItem(String value);
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
