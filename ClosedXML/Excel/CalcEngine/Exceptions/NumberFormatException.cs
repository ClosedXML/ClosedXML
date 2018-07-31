using System;
using System.Collections.Generic;
using System.Text;

namespace ClosedXML.Excel.CalcEngine.Exceptions
{
    /// <summary>
    /// This error can be caused by an invalid format of a number in an argument
    /// of an Excel function or a formula
    /// </summary>
    /// <seealso cref="ClosedXML.Excel.CalcEngine.Exceptions.CalcEngineException" />
    public class NumberFormatException : CalcEngineException
    {
        internal NumberFormatException()
            : base()
        { }

        internal NumberFormatException(string message)
            : base(message)
        { }

        internal NumberFormatException(string message, Exception innerException)
            : base(message, innerException)
        { }
    }
}
