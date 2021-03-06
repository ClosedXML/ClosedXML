using System;

namespace ClosedXML.Excel.CalcEngine.Exceptions
{
    /// <summary>
    /// This error value appears when you incorrectly type the range
    /// name, refer to a deleted range name, or forget to put quotation
    /// marks around a text string in a formula.
    /// Corresponds to the #NAME? error in Excel
    /// </summary>
    /// <seealso cref="System.ApplicationException" />
    internal class NameNotRecognizedException : CalcEngineException
    {
        internal NameNotRecognizedException()
            : base()
        { }

        internal NameNotRecognizedException(string message)
            : base(message)
        { }

        internal NameNotRecognizedException(string message, Exception innerException)
            : base(message, innerException)
        { }

    }
}
