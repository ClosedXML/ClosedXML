#nullable disable

using System;

namespace ClosedXML.Excel.CalcEngine.Exceptions
{
    /// <summary>
    /// An exception to propagate error from legacy expression function.
    /// </summary>
    internal class FormulaErrorException : Exception
    {
        public XLError Error { get; }

        public FormulaErrorException(XLError error)
        {
            Error = error;
        }
    }
}
