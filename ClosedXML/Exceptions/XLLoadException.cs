using System;

namespace ClosedXML.Exceptions
{
    public class XLLoadException : Exception
    {
        internal XLLoadException()
            : base()
        { }

        internal XLLoadException(string message)
            : base(message)
        { }

        internal XLLoadException(string message, Exception innerException)
            : base(message, innerException)
        { }
    }
}
