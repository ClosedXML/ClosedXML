using System;

namespace ClosedXML.Excel.CalcEngine.Exceptions
{
    internal abstract class CalcEngineException : ArgumentException
    {
        public CalcEngineException()
            : base()
        { }
        public CalcEngineException(string message)
            : base(message)
        { }

        public CalcEngineException(string message, Exception innerException)
            : base(message, innerException)
        { }
    }
}
