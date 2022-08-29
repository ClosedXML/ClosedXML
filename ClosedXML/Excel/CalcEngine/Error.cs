namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// A formula error.
    /// </summary>
    public enum Error
    {
        /// <summary>
        /// #REF!
        /// </summary>
        /// <remarks>When unable to find a sheet or a cell.</remarks>
        CellReference,

        /// <summary>
        /// #VALUE!
        /// </summary>
        /// <remarks>Intended to indicate when an incompatible type argument is passed to a function, or an incompatible type operand is used with an operator.</remarks>
        CellValue,

        /// <summary>
        /// #DIV/0!
        /// </summary>
        DivisionByZero,

        /// <summary>
        /// #NAME?
        /// </summary>
        /// <remarks>When unable to find a named range (but not a sheet!)</remarks>
        NameNotRecognized,

        /// <summary>
        /// #N/A
        /// </summary>
        NoValueAvailable,

        /// <summary>
        /// #NULL!
        /// </summary>
        NullValue,

        /// <summary>
        /// #NUM!
        /// </summary>
        NumberInvalid
    }
}
