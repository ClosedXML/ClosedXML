using System;

namespace ClosedXML.IO;

/// <summary>
/// An exception thrown from parser when there is a problem with data in XML.
/// The exception messages are rather generic and not very helpful, but they
/// aren't supposed to be. If this exception is thrown, there is either
/// a problem with producer of a workbook or ClosedXML. Both should do
/// investigation based on a the file causing an error.
/// </summary>
public class PartStructureException : Exception
{
    private PartStructureException(string message, XmlTreeReader? reader = null, Exception? innerException = null)
        : base(BuildMessage(message, reader), innerException)
    {
    }

    /// <summary>
    /// Create a new exception with info that some element that should be present in a workbook
    /// is missing.
    /// </summary>
    public static Exception ExpectedElementNotFound()
    {
        return new PartStructureException($"The structure of XML expected a certain kind of element, but it isn't there.");
    }

    /// <summary>
    /// Create a new exception with info that some element that should be present in a workbook
    /// is missing.
    /// </summary>
    /// <param name="missingElementDesc">optional info about what element is missing.</param>
    /// <param name="reader">XML reader at the position of the error.</param>
    public static Exception ExpectedElementNotFound(string missingElementDesc, XmlTreeReader? reader = null)
    {
        return new PartStructureException($"The structure of XML expected a certain kind of element, but it isn't there ({missingElementDesc}).", reader);
    }

    /// <summary>
    /// XML should be one of several elements, but none of them is found. Instead, there is an
    /// unexpected element.
    /// </summary>
    public static Exception ExpectedChoiceElementNotFound(XmlTreeReader reader)
    {
        return new PartStructureException($"The structure of XML expected an element from choice of several, but found {reader.ElementName} instead.", reader);
    }

    /// <summary>
    /// XML should contain a certain number of elements to be valid, but expected number of
    /// elements is different from expected one.
    /// </summary>
    public static Exception IncorrectElementsCount()
    {
        return new PartStructureException("There is a problem with element structure in XML, the number of elements found is not what was expected.");
    }

    /// <summary>
    /// XML element should contain some children or attributes and it doesn't.
    /// </summary>
    /// <param name="elementName">Name of the element.</param>
    public static Exception IncorrectElementFormat(string elementName)
    {
        return new PartStructureException($"The element '{elementName}' doesn't have or misses child elements/attributes that are required by constrains of the workbook.");
    }

    /// <summary>
    /// XML shouldn't contain an element at that point, but there is an element. This is more
    /// generic version of <see cref="ExpectedChoiceElementNotFound"/>. That one is where there are
    /// some choices and one should be there, this is generic error that element was found where it
    /// shouldn't be.
    /// </summary>
    /// <param name="elementName">Name of found element.</param>
    public static Exception UnexpectedElementFound(string elementName)
    {
        return new PartStructureException($"At this point, there shouldn't be element '{elementName}', but it is present.");
    }

    /// <summary>
    /// XML must contain a specific element, but doesn't.
    /// </summary>
    /// <param name="elementName">Name of element that should be there.</param>
    public static Exception RequiredElementIsMissing(string elementName)
    {
        return new PartStructureException($"The XML schema requires an element '{elementName}', but is is not present.");
    }

    /// <inheritdoc cref="MissingAttribute(string,XmlTreeReader)"/>
    public static Exception MissingAttribute()
    {
        return new PartStructureException("XML doesn't contain a required attribute.");
    }

    /// <inheritdoc cref="MissingAttribute(string,XmlTreeReader)"/>
    public static Exception MissingAttribute(string attributeName)
    {
        return new PartStructureException($"XML doesn't contain a required attribute '{attributeName}'.");
    }

    /// <summary>
    /// XML element must contain an attribute (generally because other element in XML), but that
    /// attribute is not in the element.
    /// </summary>
    /// <param name="attributeName">Name of attribute.</param>
    /// <param name="reader">Reader to provide info about place where error happened.</param>
    public static Exception MissingAttribute(string attributeName, XmlTreeReader reader)
    {
        var message = $"XML doesn't contain a required attribute '{attributeName}'.";
        return new PartStructureException(message, reader);
    }

    /// <inheritdoc cref="InvalidAttributeFormat(string)"/>
    public static Exception InvalidAttributeFormat()
    {
        return new PartStructureException("The attribute has a value in an incorrect format.");
    }

    /// <summary>
    /// Attribute value should have some kind of format (e.g. number or an enum value) and it
    /// doesn't.
    /// </summary>
    /// <param name="attributeValue">Value of the attribute.</param>
    public static Exception InvalidAttributeFormat(string attributeValue)
    {
        return new PartStructureException($"The attribute has a value ('{attributeValue}') in an incorrect format.");
    }

    /// <summary>
    /// Attribute value should have some kind of format (e.g. number or an enum value) and it
    /// doesn't.
    /// </summary>
    /// <param name="attributeName">Name of the attribute.</param>
    /// <param name="attributeValue">Value of the attribute.</param>
    /// <param name="reader">Reader to provide info about place where error happened.</param>
    /// <param name="innerException">Exception that cause the error.</param>
    public static Exception InvalidAttributeFormat(string attributeName, string attributeValue, XmlTreeReader reader, Exception? innerException = null)
    {
        return new PartStructureException($"The attribute '{attributeName}' contains a value '{attributeValue}' that doesn't match expected format.", reader, innerException);
    }

    /// <inheritdoc cref="InvalidAttributeValue(string)"/>
    public static Exception InvalidAttributeValue()
    {
        return new PartStructureException("The value of attribute doesn't make sense with the rest of data of a workbook (e.g. reference that doesn't exist).");
    }

    /// <summary>
    /// The attribute value doesn't make sense when taken in context of whole XML document. That is
    /// different from <see cref="InvalidAttributeFormat(string)"/>, format is a syntactic problem, this
    /// is a semantic problem.
    /// </summary>
    /// <param name="attributeValue">The attribute value, not a name.</param>
    public static Exception InvalidAttributeValue(string attributeValue)
    {
        return new PartStructureException($"The value of attribute '{attributeValue}' is not valid value for the attribute.");
    }


    private static string BuildMessage(string message, XmlTreeReader? reader)
    {
        if (reader is not null && reader.TryGetLineInfo(out var lineInfo))
        {
            message += $" Line:{lineInfo.LineNumber}, Position:{lineInfo.LinePosition}.";
        }

        return message;
    }
}
