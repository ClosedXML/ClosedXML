namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// A formula error.
    /// </summary>
    public enum XLError
    {
        /// <summary>
        /// <c>#REF!</c> - a formula refers to a cell that's not valid.
        /// </summary>
        /// <example>When unable to find a sheet or a cell.</example>
        CellReference,

        /// <summary>
        /// <c>#VALUE!</c> - Intended to indicate when an incompatible type argument is passed to a function, or an incompatible type operand is used with an operator.
        /// </summary>
        /// <example>Passing a non-number text to a function that requires a number, trying to get an area from non-contiguous reference. Creating an area from different sheets <c>Sheet1!A1:Sheet2!A2</c></example>
        IncompatibleValue,

        /// <summary>
        /// <c>#DIV/0!</c> - Intended to indicate when any number (including zero) or any error code is divided by zero.
        /// </summary>
        DivisionByZero,

        /// <summary>
        /// <c>#NAME?</c> - Intended to indicate when what looks like a name is used, but no such name has been defined.
        /// </summary>
        /// <remarks>Only for named ranges, not sheets.</remarks>
        /// <example><c>TestRange*10</c> when the named range doesn't exist will result in an error.</example>
        NameNotRecognized,

        /// <summary>
        /// <c>#N/A</c> - Intended to indicate when a designated value is not available.
        /// </summary>
        /// <example>The value is used for extra cells of an array formula that is applied on an array of a smaller size that the array formula.</example>
        NoValueAvailable,

        /// <summary>
        /// <c>#NULL!</c> - Intended to indicate when two areas are required to intersect, but do not.
        /// </summary>
        /// <remarks>The space is an intersection operator.</remarks>
        /// <example><c>SUM(B1 C1)</c> tries to intersect <c>B1:B1</c> area and <c>C1:C1</c> area, but since there are no intersecting cells, the result is <c>#NULL</c>.</example>
        NullValue,

        /// <summary>
        /// <c>#NUM!</c> - Intended to indicate when an argument to a function has a compatible type, but has a value that is outside the domain over which that function is defined.
        /// </summary>
        /// <remarks>This is known as a domain error.</remarks>
        /// <example>ASIN(10) - the ASIN accepts only argument -1..1 (an output of SIN), so the resulting value is <c>#NUM!</c>.</example>
        NumberInvalid
    }
}
