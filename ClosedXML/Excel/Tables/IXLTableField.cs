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
        IXLRangeColumn Column { get; }
        Int32 Index { get; }
        String Name { get; set; }
        String TotalsRowFormulaA1 { get; set; }
        String TotalsRowFormulaR1C1 { get; set; }
        XLTotalsRowFunction TotalsRowFunction { get; set; }
        String TotalsRowLabel { get; set; }

        void Delete();

        /// <summary>
        /// Determines whether all cells this table field have a consistent data type.
        /// </summary>
        Boolean IsConsistentDataType();
    }
}
