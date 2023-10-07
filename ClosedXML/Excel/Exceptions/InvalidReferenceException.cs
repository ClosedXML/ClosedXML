using System;

namespace ClosedXML.Excel.Exceptions
{
    /// <summary>
    /// A reference to the data in a worksheet is not valid. E.g. sheet with
    /// specific name doesn't exist, name doesn't exist.
    /// </summary>
    public class InvalidReferenceException : Exception
    {
        public InvalidReferenceException() : base("Reference to the data is not valid.")
        {
        }

        public InvalidReferenceException(string message)
            : base(message)
        {
        }

        public InvalidReferenceException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}
