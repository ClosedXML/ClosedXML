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
        string SourceName { get; }
        string CustomName { get; set; }

        IXLPivotValueFormat NumberFormat { get; }

        XLPivotSummary SummaryFormula { get; set; }
        XLPivotCalculation Calculation { get; set; }
        string BaseField { get; set; }
        string BaseItem { get; set; }
        XLPivotCalculationItem CalculationItem { get; set; }

        IXLPivotValue SetSummaryFormula(XLPivotSummary value);
        IXLPivotValue SetCalculation(XLPivotCalculation value);
        IXLPivotValue SetBaseField(string value);
        IXLPivotValue SetBaseItem(string value);
        IXLPivotValue SetCalculationItem(XLPivotCalculationItem value);


        IXLPivotValue ShowAsNormal();
        IXLPivotValueCombination ShowAsDifferenceFrom(string fieldSourceName);
        IXLPivotValueCombination ShowAsPercentageFrom(string fieldSourceName);
        IXLPivotValueCombination ShowAsPercentageDifferenceFrom(string fieldSourceName);
        IXLPivotValue ShowAsRunningTotalIn(string fieldSourceName);
        IXLPivotValue ShowAsPercentageOfRow();
        IXLPivotValue ShowAsPercentageOfColumn();
        IXLPivotValue ShowAsPercentageOfTotal();
        IXLPivotValue ShowAsIndex();

    }
}
