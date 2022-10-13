using System;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// Function flags that indicate what does function do. It is used by CalcEngine for calculation
    /// chain and formula execution.
    /// </summary>
    [Flags]
    internal enum FunctionFlags
    {
        /// <summary>
        /// Function that takes an input and returns an output. It is designed for a single value arguments.
        /// If scalar function is used for array formula or dynamic array formula, the function is called for each element separately.
        /// </summary>
        Scalar = 0,

        /// <summary>
        /// Non-scalar function. At least one of arguments of the function accepts a range. It means that
        /// implicit intersection works differently.
        /// </summary>
        Range = 1,

        /// <summary>
        /// Function has side effects, e.g. it changes something.
        /// </summary>
        /// <example>HYPERLINK</example>
        SideEffect = 2
    }
}
