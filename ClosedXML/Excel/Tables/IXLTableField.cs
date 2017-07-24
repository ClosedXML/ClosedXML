using System;

namespace ClosedXML.Excel
{
    public enum XLTotalsRowFunction
    {
        None,
        Sum,
        Minimum,
        Maximum,
        Average,
        Count,
        CountNumbers,
        StandardDeviation,
        Variance,
        Custom
    }

    public interface IXLTableField
    {
        Int32 Index { get; }
        String Name { get; set; }
        String TotalsRowLabel { get; set; }
        String TotalsRowFormulaA1 { get; set; }
        String TotalsRowFormulaR1C1 { get; set; }
        XLTotalsRowFunction TotalsRowFunction { get; set; }
    }
}
