using System;

namespace ClosedXML.Excel.IO
{
    /// <summary>
    /// An exception thrown from parser when there is a problem with data in XML.
    /// The exception messages are rather generic and not very helpful, but they
    /// aren't supposed to be. If this exception is thrown, there is either
    /// a problem with producer of a workbook or ClosedXML. Both should do
    /// investigation based on a the file causing an error.
    /// </summary>
    internal class PartStructureException : Exception
    {
        private PartStructureException(string message, string? detail = null)
            : base(detail is null ? message : message[..^1] + " (" + detail + ").")
        {
        }

        /// <summary>
        /// Create a new exception with info that some element that should be present in a workbook
        /// is missing.
        /// </summary>
        /// <param name="missingElementDesc">optional info about what element is missing.</param>
        internal static Exception ExpectedElementNotFound(string? missingElementDesc = null)
        {
            return new PartStructureException("The structure of XML expected a certain kind of element, but it isn't there.", missingElementDesc);
        }

        internal static Exception IncorrectElementsCount()
        {
            return new PartStructureException("There is a problem with element structure in XML, the number of elements found is not what was expected.");
        }

        internal static Exception MissingAttribute()
        {
            return new PartStructureException("XML doesn't contain a required attribute.");
        }

        internal static Exception IncorrectAttributeFormat()
        {
            return new PartStructureException("The attribute has a value in an incorrect format.");
        }

        internal static Exception IncorrectAttributeValue()
        {
            return new PartStructureException("The value of attribute doesn't make sense with the rest of data of a workbook (e.g. reference that doesn't exist).");
        }
    }
}
