using System;

namespace ClosedXML.Excel.CalcEngine.Exceptions
{
    /// <summary>
    /// Because a space indicates an intersection, this error will
    /// occur if you insert a space instead of a comma(the union operator)
    /// between ranges used in function arguments.
    /// Corresponds to the #NULL! error in Excel
    /// </summary>
    /// <seealso cref="ClosedXML.Excel.CalcEngine.Exceptions.CalcEngineException" />
    internal class NullValueException : CalcEngineException
    {
        internal NullValueException()
            : base()
        { }

        internal NullValueException(string message)
            : base(message)
        { }

        internal NullValueException(string message, Exception innerException)
            : base(message, innerException)
        { }

    }
}
