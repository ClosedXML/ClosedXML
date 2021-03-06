using System;

namespace ClosedXML.Excel.CalcEngine.Exceptions
{
    /// <summary>
    /// This error is most often the result of specifying a
    /// mathematical operation with one or more cells that contain
    /// text.
    /// Corresponds to the #VALUE! error in Excel
    /// </summary>
    /// <seealso cref="ClosedXML.Excel.CalcEngine.Exceptions.CalcEngineException" />
    internal class CellValueException : CalcEngineException
    {
        internal CellValueException()
            : base()
        { }

        internal CellValueException(string message)
            : base(message)
        { }

        internal CellValueException(string message, Exception innerException)
            : base(message, innerException)
        { }
    }
}
