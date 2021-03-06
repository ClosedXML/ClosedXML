using System;

namespace ClosedXML.Excel.CalcEngine.Exceptions
{
    internal abstract class CalcEngineException : ArgumentException
    {
        protected CalcEngineException()
            : base()
        { }
        protected CalcEngineException(string message)
            : base(message)
        { }

        protected CalcEngineException(string message, Exception innerException)
            : base(message, innerException)
        { }
    }
}
