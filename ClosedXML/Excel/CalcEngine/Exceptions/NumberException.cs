using System;

namespace ClosedXML.Excel.CalcEngine.Exceptions
{
    /// <summary>
    /// This error can be caused by an invalid argument in an Excel
    /// function or a formula that produces a number too large or too small
    /// to be represented in the worksheet.
    /// Corresponds to the #NUM! error in Excel
    /// </summary>
    /// <seealso cref="ClosedXML.Excel.CalcEngine.Exceptions.CalcEngineException" />
    internal class NumberException : CalcEngineException
    {
        internal NumberException()
            : base()
        { }

        internal NumberException(string message)
            : base(message)
        { }

        internal NumberException(string message, Exception innerException)
            : base(message, innerException)
        { }

    }
}
