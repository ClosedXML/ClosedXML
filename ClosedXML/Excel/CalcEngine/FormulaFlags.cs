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
        None = 0,

        /// <summary>
        /// Formula contains a function whose value can be different each time the function is called, even if arguments and workbook are same.
        /// </summary>
        /// <example><c>TODAY</c>, <c>RAND</c></example>
        Volatile = 1,

        /// <summary>
        /// Formula that has a side effects beside returning the value.
        /// </summary>
        /// <example><c>HYPERLINK</c> changes the content of a cell in a workbook.</example>
        SideEffect = 2,

        /// <summary>
        /// Formula contains a reference to the <c>SUBTOTAL</c> function.
        /// </summary>
        /// <remarks>Performance optimization, so a formula with SUBTOTAL doesn't have to check each dependent cell each time it is evaluated.</remarks>
        HasSubtotal = 4
    }
}
