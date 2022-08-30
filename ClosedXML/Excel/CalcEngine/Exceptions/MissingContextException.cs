using System;

namespace ClosedXML.Excel.CalcEngine.Exceptions
{
    /// <summary>
    /// Evaluation of the formula needs an information that wasn't available. That can happen if the formula
    /// is evaluated from methods like <see cref="XLWorkbook.Evaluate(string)"/>. Causes vary, e.g. implicit intersection
    /// needs an address of the formula cell.
    /// </summary>
    internal class MissingContextException : InvalidOperationException
    {
    }
}
