using System;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// Flags that contain some useful information about a formula (mostly derived from flags in functions).
    /// </summary>
    [Flags]
    internal enum FormulaFlags : byte
    {
        /// <summary>
        /// Basic formula that takes an input and returns output that is determined solely by the input. No side effects.
        /// </summary>
        None = 0
    }
}
