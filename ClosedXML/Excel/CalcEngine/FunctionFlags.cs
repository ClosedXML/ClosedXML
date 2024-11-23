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
        SideEffect = 2,

        /// <summary>
        /// Function returns array. Functions without this flag return a scalar value.
        /// CalcEngine treats such functions differently for array formulas.
        /// </summary>
        ReturnsArray = 4,

        /// <summary>
        /// Function is not deterministic.
        /// </summary>
        /// <example>RAND(), DATE()</example>
        Volatile = 8,

        /// <summary>
        /// The function is a future function (i.e. functions not present in Excel 2007). Future
        /// functions are displayed to the user with a name (e.g <c>SEC</c>), but are actually
        /// stored in the workbook with a prefix <c>_xlfn</c> (e.g. <c>_xlfn.SEC</c>).
        /// The prefix is there for backwards compatibility, to not clash with user defined
        /// functions and other such reasons. See [MS-XLSX] 2.3.3 for complete list.
        /// </summary>
        Future = 16
    }
}
