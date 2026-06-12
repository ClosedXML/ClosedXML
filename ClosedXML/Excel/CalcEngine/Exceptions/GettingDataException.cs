using System;

namespace ClosedXML.Excel.CalcEngine.Exceptions
{
    /// <summary>
    /// Exception that happens when formula in a cell depends on other cells,
    /// but the supporting formulas are still dirty.
    /// </summary>
    internal class GettingDataException : Exception
    {
        public GettingDataException(SheetPoint point)
        {
            Point = point;
        }

        public SheetPoint Point { get; }
    }
}
