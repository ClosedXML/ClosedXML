using System;

namespace ClosedXML.Excel.CalcEngine.Exceptions
{
    /// <summary>
    /// Technically, this is not an error value but a special value
    /// that you can manually enter into a cell to indicate that you
    /// don’t yet have a necessary value.
    /// Corresponds to the #N/A error in Excel.
    /// </summary>
    /// <seealso cref="System.ApplicationException" />
    internal class NoValueAvailableException : CalcEngineException
    {
        internal NoValueAvailableException()
            : base()
        { }

        internal NoValueAvailableException(string message)
            : base(message)
        { }

        internal NoValueAvailableException(string message, Exception innerException)
            : base(message, innerException)
        { }

    }
}
