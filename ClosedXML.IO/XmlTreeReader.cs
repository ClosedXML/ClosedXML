using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Xml;
using static ClosedXML.IO.XmlTreeNodeType;

namespace ClosedXML.IO;

/// <summary>
/// <para>
/// Reader that expects that XML document consists of elements in a tree-like fashion. XML
/// shouldn't be mixed, text should be only in the leaves (e.g. <c>&lt;f&gt;ABS(A1)*A2&lt;/f&gt;</c>).
/// </para>
/// <para>
/// The schema of XML should be mostly by elements, with choices and sequences.
/// </para>
/// <para>
/// In case of some specialities or mixed content, use <c>XDocument.Load(_reader.ReadSubtree())</c>
/// and parse the result.
/// </para>
/// <para>
/// All <c>Get*</c> methods read values from attributes of current element.
/// </para>
/// <para>
/// The reader is always at either start element or end element. Any API that moves the reader will
/// end on either start or end element. The empty elements (e.g. &lt;br/&gt;) behave same way as
/// non-empty elements (that is different from <see cref="XmlReader"/> that checks
/// <see cref="XmlReader.IsEmptyElement"/>).
/// The reader element is used for one of two purposes:
/// <list type="bullet">
/// <item>
/// Element is being processed, i.e. parser logic has correctly identified the element (by
/// name) and will use parser logic to extract data from element (mostly by reading attributes,
/// potentially content if it is a leaf).
/// </item>
/// <item>
/// Element is used as a lookahead. The parsing logic is using it to determine how to parse the rest
/// of document. Example:
/// <example>Stylesheet parser has processed element <![CDATA[<name/>]]> from <![CDATA[<font><name val="Arial"/></font>]]>).
/// It reads next element <![CDATA[</font>]]> as a lookahead. The parsing logic now has to
/// determine what to do. Font could have additional properties (e.g. <![CDATA[<b/>]]>) or the
/// schema could end. Parser will use <c>reader.TryOpen("b")</c> to check if there is a bold
/// element. If there isn't, it uses <c>reader.TryClose("font")</c> to check that <c>font</c>
/// should close. If neither is true, XML doesn't match expected schema and parser will likely
/// throw an exception.
/// </example>
/// </item>
/// </list>
/// </para>
/// </summary>
public sealed class XmlTreeReader : IDisposable
{
    private static readonly XmlReaderSettings Settings = new()
    {
        IgnoreWhitespace = true,
        IgnoreComments = true,
        IgnoreProcessingInstructions = true,
        Async = false,
        DtdProcessing = DtdProcessing.Prohibit,
        CloseInput = true
    };

    /// <summary>
    /// The reader that holds the current element. The current node must always be either
    /// <see cref="OpenElement"/> or <see cref="CloseElement"/> after an API method call.
    /// </summary>
    private readonly IXmlReader _reader;

    private readonly IEnumMapper _enumMapper;

    /// <summary>
    /// What is current state of parser:
    /// <list type="bullet">
    /// <item>
    ///   <term>false</term>
    ///   <description>
    ///   Current element is being processed. we can get value of attributes.
    ///   </description>
    /// </item>
    /// <item>
    ///   <term>true</term>
    ///   <description>
    ///   Current element is not being processed. The only thing we are interested in is a name and
    ///   open/close. We are using it to determine how to parse the remainder of the file.
    ///   Trying to get attribute value will throw.
    ///   </description>
    /// </item>
    /// </list>
    /// </summary>
    private bool _inLookup;

    private readonly List<string> _context = new();

    public XmlTreeReader(Stream stream, IEnumMapper enumMapper, bool strictAttributeParsing)
    {
        _reader = new MceXmlReader(XmlReader.Create(stream, Settings), new MceSettings
        {
            SignalMismatch = info => throw PartStructureException.MceError(info.LineInfo, "Mismatch between consuming application capability and document requirements.")
        });
        _enumMapper = enumMapper;
        StrictAttributeParsing = strictAttributeParsing;
    }

    /// <summary>
    /// Returns names of elements from currently processed element to root. The <c>On*Parsed</c>
    /// hook is called after processed element is closed. Therefore when the context is inspected
    /// in the <c>On*Parsed</c> hook, it doesn't contain the name of element that was just processed.
    /// </summary>
    public IReadOnlyList<string> Context => _context;

    /// <summary>
    /// Should attributes with values outside of their simple type be treated as errors and throw
    /// an exception or be treated as missing and just return null? Excel generally ignores
    /// unparseable attributes.
    /// </summary>
    public bool StrictAttributeParsing { get; }

    /// <summary>
    /// Get the local name of current element (lookup/processing).
    /// </summary>
    internal string ElementName => _reader.LocalName;

    internal LineInfo LineInfo => _reader.LineInfo;

    /// <summary>
    /// Read next element. Check lookup element is <paramref name="localName"/>. If it is, open the
    /// element and return true. Otherwise, return false (element doesn't change).
    /// </summary>
    public bool TryOpen(string localName, string namespaceUri)
    {
        SwitchToLookup();
        Debug.Assert(_reader.NodeType is OpenElement or CloseElement);

        if (_reader.NodeType == OpenElement && _reader.LocalName == localName && _reader.NamespaceUri == namespaceUri)
        {
            // Element has been opened, so it should be processed.
            SwitchToProcessing();
            _context.Add(localName);
            return true;
        }

        return false;
    }

    public bool TryClose(string localName, string namespaceUri)
    {
        SwitchToLookup();
        Debug.Assert(_reader.NodeType is OpenElement or CloseElement);

        if (_reader.NodeType == OpenElement || _reader.LocalName != localName || _reader.NamespaceUri != namespaceUri)
            return false;

        // Element has been closed, so it should be processed. Though closing elements are not
        // really processed, but we don't want to switch back to lookup. We just want to mark it
        // as "done." Be lazy, e.g. last element of a document doesn't have next element. If
        // parsing logic needs further elements, it will read them when they are needed.
        SwitchToProcessing();
        _context.RemoveAt(_context.Count - 1);
        return true;
    }

    /// <summary>
    /// Assert that we are at the element with <paramref name="localName"/>. Doesn't move anywhere.
    /// </summary>
    public void Open(string localName, string namespaceUri)
    {
        if (!TryOpen(localName, namespaceUri))
            throw PartStructureException.ExpectedElementNotFound($"Expected opening element '{localName}', but reader is currently on {_reader.NodeType} node '{_reader.LocalName}'.", this);
    }

    /// <summary>
    /// Close the next unprocessed node. If the node doesn't match the <paramref name="localName"/>,
    /// throw an exception.
    /// </summary>
    public void Close(string localName, string namespaceUri)
    {
        if (!TryClose(localName, namespaceUri))
            throw PartStructureException.ExpectedElementNotFound($"Expected closing element '{localName}', but reader is currently on {_reader.NodeType} node '{_reader.LocalName}'.", this);
    }

    /// <summary>
    /// Skip subtree that start on the current element. After subtree is read, the reader is
    /// on an ending element of a subtree in a processed state.
    /// </summary>
    /// <exception cref="InvalidOperationException">Reader isn't on opening element.</exception>
    public void Skip(string elementName)
    {
        ThrowOnNonStartElement();
        var startDepth = _reader.Depth;
        do
        {
            if (!_reader.Read())
                throw InvalidXml();
        } while (_reader.NodeType != CloseElement || _reader.Depth > startDepth);

        _inLookup = false;
    }

    public bool? GetOptionalBool(string attributeName)
    {
        ThrowOnNonStartElement();
        bool? result = null;
        if (_reader.GetAttribute(attributeName, null) is { } value)
        {
            try
            {
                result = XmlConvert.ToBoolean(value);
            }
            catch (FormatException e)
            {
                if (StrictAttributeParsing)
                    ThrowAttributeFormatException(attributeName, value, e);
            }
        }

        return result;
    }

    public int? GetOptionalInt(string attributeName)
    {
        ThrowOnNonStartElement();
        int? number = null;
        if (_reader.GetAttribute(attributeName, null) is { } value)
        {
            try
            {
                number = XmlConvert.ToInt32(value);
            }
            catch (OverflowException e)
            {
                if (StrictAttributeParsing)
                    ThrowAttributeFormatException(attributeName, value, e);
            }
            catch (FormatException e)
            {
                if (StrictAttributeParsing)
                    ThrowAttributeFormatException(attributeName, value, e);
            }
        }

        return number;
    }

    public uint? GetOptionalUInt(string attributeName)
    {
        ThrowOnNonStartElement();
        long? number = null;
        if (_reader.GetAttribute(attributeName, null) is { } value)
        {
            try
            {
                number = XmlConvert.ToInt64(value);
            }
            catch (OverflowException e)
            {
                if (StrictAttributeParsing)
                    ThrowAttributeFormatException(attributeName, value, e);
            }
            catch (FormatException e)
            {
                if (StrictAttributeParsing)
                    ThrowAttributeFormatException(attributeName, value, e);
            }

            if (number is < 0 or > uint.MaxValue)
            {
                if (StrictAttributeParsing)
                    ThrowAttributeFormatException(attributeName, value);

                number = null;
            }
        }

        return number is not null ? (uint)number : null;
    }

    public double? GetOptionalDouble(string attributeName)
    {
        ThrowOnNonStartElement();
        double? number = null;
        if (_reader.GetAttribute(attributeName, null) is { } value)
        {
            try
            {
                number = XmlConvert.ToDouble(value);
            }
            catch (OverflowException e)
            {
                if (StrictAttributeParsing)
                    ThrowAttributeFormatException(attributeName, value, e);
            }
            catch (FormatException e)
            {
                if (StrictAttributeParsing)
                    ThrowAttributeFormatException(attributeName, value, e);
            }

            if (number is not null && (double.IsNaN(number.Value) || double.IsInfinity(number.Value)))
            {
                if (StrictAttributeParsing)
                    ThrowAttributeFormatException(attributeName, value);

                number = null;
            }
        }

        return number;
    }

    /// <summary>
    /// Try to read <c>xsd:dateTime</c> from an attribute of current element.
    /// </summary>
    /// <param name="attributeName">Name of the attribute.</param>
    /// <returns>Read datetime or null if attribute is not present.</returns>
    public DateTime? GetOptionalDateTime(string attributeName)
    {
        ThrowOnNonStartElement();
        DateTime? dateTime = null;
        if (_reader.GetAttribute(attributeName, null) is { } value)
        {
            try
            {
                dateTime = XmlConvert.ToDateTime(value, XmlDateTimeSerializationMode.RoundtripKind);
            }
            catch (FormatException e)
            {
                if (StrictAttributeParsing)
                    ThrowAttributeFormatException(attributeName, value, e);
            }
        }

        return dateTime;
    }

    public string? GetOptionalString(string attributeName)
    {
        ThrowOnNonStartElement();
        return _reader.GetAttribute(attributeName, null);
    }

    public TEnum? GetOptionalEnum<TEnum>(string attributeName)
        where TEnum : struct, Enum
    {
        ThrowOnNonStartElement();
        var enumString = _reader.GetAttribute(attributeName, null);

        if (enumString is null)
        {
            return null;
        }

        if (!_enumMapper.TryGetEnum<TEnum>(enumString, out var enumValue))
        {
            if (StrictAttributeParsing)
                ThrowAttributeFormatException(attributeName, enumString);

            return null;
        }

        return enumValue;
    }

    /// <summary>
    /// Get the content of current element. Once it ends, it should be on
    /// the <see cref="XmlTreeNodeType.CloseElement"/> of the tree.
    /// </summary>
    /// <exception cref="PartStructureException">If another element is found inside the current element.</exception>>
    public string GetContent()
    {
        ThrowOnNonStartElement();

        var content = string.Empty;
        while (_reader.Read())
        {
            switch (_reader.NodeType)
            {
                case OpenElement:
                    throw PartStructureException.UnexpectedElementFound(_reader.LocalName);
                case CloseElement:
                    _inLookup = true;
                    return content;
                case Text:
                    content += _reader.Value;
                    break;
                default:
                    throw new UnreachableException();
            }
        }

        throw InvalidXml();
    }

    public void Dispose()
    {
        _reader.Dispose();
    }

    private void SwitchToProcessing()
    {
        if (_inLookup)
            _inLookup = false;
    }

    private void SwitchToLookup()
    {
        // When switching to lookup, current node and all its attributes should have already been processed.
        if (_inLookup)
            return;

        // Read next element.
        MoveToNextElement();
        _inLookup = true;
    }

    /// <summary>
    /// Move from current opening/closing element to next opening/closing element.
    /// </summary>
    private void MoveToNextElement()
    {
        while (_reader.Read())
        {
            if (_reader.NodeType is OpenElement or CloseElement)
            {
                return;
            }
        }

        throw InvalidXml();
    }

    private void ThrowOnNonStartElement()
    {
        if (_reader.NodeType != OpenElement || _inLookup)
            throw new InvalidOperationException("To read content/attribute, the reader must be on start element and in processing state.");
    }

    private void ThrowAttributeFormatException(string attributeName, string attributeValue, Exception? exception = null)
    {
        throw PartStructureException.InvalidAttributeFormat(attributeName, attributeValue, this, exception);
    }

    private Exception InvalidXml()
    {
        // This should never happen. The underlaying reader should throw when it detects
        // invalid XML (no root, unpaired elements, multiple elements in root)
        return new UnreachableException($"{_reader.LineInfo}: Invalid XML.");
    }
}
