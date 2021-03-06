using System;

namespace ClosedXML.Excel.CalcEngine.Exceptions
{
    /// <summary>
    /// This error occurs when you delete a cell referred to in the
    /// formula or if you paste cells over the ones referred to in the
    /// formula.
    /// Corresponds to the #REF! error in Excel
    /// </summary>
    /// <seealso cref="ClosedXML.Excel.CalcEngine.Exceptions.CalcEngineException" />
    internal class CellReferenceException : CalcEngineException
    {
        internal CellReferenceException()
            : base()
        { }

        internal CellReferenceException(string message)
            : base(message)
        { }

        internal CellReferenceException(string message, Exception innerException)
            : base(message, innerException)
        { }

    }
}
