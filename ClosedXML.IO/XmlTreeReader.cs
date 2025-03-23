using System;
using System.Diagnostics.CodeAnalysis;
using System.Xml;

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
/// <remarks>
/// <para>
/// When adding API, use <em>tell, don't ask principle</em>. Asking generally inherently requires
/// allocation of a new string, but passing a string to a method doesn't.
/// </para>
/// <para>Example:
/// <example>
/// <c>reader.IsStartElement("font")</c> vs <c>reader.LocalName == "font"</c>. The former methods
/// re-uses interned string. If reader is optimized, it can check everything on stack, without any
/// new allocations. The second comparison basically requires a new string allocation for the
/// <see cref="XmlReader.LocalName"/> getter.
/// </example>
/// </para>
/// <para>
/// Another example would be <see cref="XmlReader.ReadContentAsBoolean"/> instead of getting string
/// and parsing it ourselves. Allocations do matter when parsing hundreds of MBs.
/// </para>
/// </remarks>
public sealed class XmlTreeReader : IDisposable
{
    /// <summary>
    /// The XmlReader that holds current element. The current node should always be either
    /// <see cref="XmlNodeType.Element"/> or <see cref="XmlNodeType.EndElement"/>.
    /// </summary>
    private readonly XmlReader _reader;

    private readonly IEnumMapper _enumMapper;

    /// <summary>
    /// <para>
    /// An abstraction to deal with empty elements. If current element is an empty element
    /// (regardless of whether in processing or lookup mode), this property determines if
    /// the element is interpreted as starting element or ending element.
    /// </para>
    /// <para>
    /// The property is set for every element to make everything easier.
    /// </para>
    /// </summary>
    private bool _isStart = true;

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
    private bool _inLookup = true;

    public XmlTreeReader(XmlReader reader, IEnumMapper enumMapper)
    {
        _reader = reader;
        _enumMapper = enumMapper;
    }

    /// <summary>
    /// Get name of current element (lookup/processing). It includes an alias for ns.
    /// </summary>
    internal string ElementName => _reader.Name;

    /// <summary>
    /// Read next element. Check lookup element is <paramref name="localName"/>. If it is, open the
    /// element and return true. Otherwise, return false (element doesn't change).
    /// </summary>
    public bool TryOpen(string localName, string namespaceUri)
    {
        ThrowWhenReaderNotOnElement();
        SwitchToLookup();

        if (_isStart && _reader.LocalName == localName && _reader.NamespaceURI == namespaceUri)
        {
            // Element has been opened, so it should be processed.
            SwitchToProcessing();
            return true;
        }

        return false;
    }

    // Throws when it is on closing elements of incorrect type
    public bool TryClose(string localName, string namespaceUri)
    {
        ThrowWhenReaderNotOnElement();
        SwitchToLookup();

        if (_isStart || _reader.LocalName != localName || _reader.NamespaceURI != namespaceUri)
            return false;

        // Element has been closed, so it should be processed. Though closing elements are not
        // really processed, but we don't want to switch back to lookup. We just want to mark it
        // as "done." Be lazy, e.g. last element of a document doesn't have next element. If
        // parsing logic needs further elements, it will read them when they are needed.
        SwitchToProcessing();
        return true;
    }

    /// <summary>
    /// Assert that we are at the element with <paramref name="localName"/>. Doesn't move anywhere.
    /// </summary>
    public void Open(string localName, string namespaceUri)
    {
        if (!TryOpen(localName, namespaceUri))
            throw PartStructureException.ExpectedElementNotFound($"Expected opening element '{localName}', but reader is currently on {(_isStart ? "opening" : "closing")} '{_reader.Name}'.");
    }

    /// <summary>
    /// Close the next unprocessed node. If the node doesn't match the <paramref name="localName"/>,
    /// throw an exception.
    /// </summary>
    public void Close(string localName, string namespaceUri)
    {
        if (!TryClose(localName, namespaceUri))
            throw PartStructureException.ExpectedElementNotFound($"Expected closing element '{localName}', but reader is currently on {(_isStart ? "opening" : "closing")} '{_reader.Name}'.");
    }

    /// <summary>
    /// Skip subtree that start on the current element. After subtree is read, the reader is
    /// on an ending element of a subtree in a processed state.
    /// </summary>
    /// <exception cref="InvalidOperationException">Reader isn't on opening element.</exception>
    public void Skip()
    {
        ThrowOnNonStartElement();
        var startDepth = _reader.Depth;
        do
        {
            MoveToNextElement();
        } while (_isStart || _reader.Depth > startDepth);

        _inLookup = false;
    }

    /// <summary>
    /// Read the content of current element. Ends in a lookup state on the end element.
    /// </summary>
    /// <exception cref="PartStructureException">The content contains elements.</exception>
    public string GetContent()
    {
        ThrowOnNonStartElement();
        if (_reader.IsEmptyElement)
        {
            _inLookup = true;
            _isStart = false;
            return string.Empty;
        }

        // ReadElementContentAsString reads beyond closing element. Make your own reader.
        var value = string.Empty;
        while (ReadNode() is { } nodeType && nodeType != XmlNodeType.EndElement)
        {
            // All unspecified nodes should be skipped. It is either comments, processing
            // instructions or something that shouldn't ever happen (e.g. attribute).
            switch (nodeType)
            {
                case XmlNodeType.Text:
                case XmlNodeType.Whitespace:
                case XmlNodeType.SignificantWhitespace:
                case XmlNodeType.CDATA:
                    if (value.Length == 0)
                        value = _reader.Value;
                    else
                        value += _reader.Value;
                    break;
                case XmlNodeType.EntityReference:
                    value += _reader.Name; // Does it even work? I have no idea how to get this node.
                    break;
                case XmlNodeType.Element:
                    throw PartStructureException.UnexpectedElementFound(_reader.LocalName); // No child elements allowed
            }
        }

        _inLookup = true;
        _isStart = false;
        return value;
    }

    public bool? GetOptionalBool(string attributeName)
    {
        ThrowOnNonStartElement();
        bool? result = _reader.MoveToAttribute(attributeName) ? _reader.ReadContentAsBoolean() : null;
        _reader.MoveToElement();
        return result;
    }

    public int? GetOptionalInt(string attributeName)
    {
        ThrowOnNonStartElement();
        _reader.MoveToAttribute(attributeName);
        int? number = _reader.MoveToAttribute(attributeName) ? _reader.ReadContentAsInt() : null;
        _reader.MoveToElement();
        return number;
    }

    public uint? GetOptionalUint(string attributeName)
    {
        ThrowOnNonStartElement();
        long? number = _reader.MoveToAttribute(attributeName) ? _reader.ReadContentAsLong() : null;
        if (number is < 0 or > uint.MaxValue)
            throw PartStructureException.InvalidAttributeFormat(_reader.ReadContentAsString());

        _reader.MoveToElement();
        return number is not null ? (uint)number : null;
    }

    public double? GetOptionalDouble(string attributeName)
    {
        ThrowOnNonStartElement();
        double? number = _reader.MoveToAttribute(attributeName) ? _reader.ReadContentAsDouble() : null;
        _reader.MoveToElement();
        return number;
    }

    public string? GetOptionalString(string attributeName)
    {
        ThrowOnNonStartElement();
        return _reader.GetAttribute(attributeName);
    }

    public TEnum? GetOptionalEnum<TEnum>(string attributeName)
        where TEnum : struct, Enum
    {
        ThrowOnNonStartElement();
        var enumString = _reader.MoveToAttribute(attributeName) ? _reader.ReadContentAsString() : null;
        _reader.MoveToElement();

        if (enumString is null)
            return null;

        if (!_enumMapper.TryGetEnum<TEnum>(enumString, out var enumValue))
            throw PartStructureException.InvalidAttributeFormat(enumString);

        return enumValue;
    }

    public void Dispose()
    {
        _reader.Dispose();
    }

    internal bool TryGetLineInfo([NotNullWhen(true)] out IXmlLineInfo? lineInfo)
    {
        if (_reader is IXmlLineInfo readerInfo && readerInfo.HasLineInfo())
        {
            lineInfo = readerInfo;
            return true;
        }

        lineInfo = null;
        return false;
    }

    private void SwitchToProcessing()
    {
        if (_inLookup)
            _inLookup = false;
    }

    private void SwitchToLookup()
    {
        ThrowWhenReaderNotOnElement();

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
        if (_isStart && _reader.IsEmptyElement)
        {
            _isStart = false;
            return;
        }

        while (ReadNode() is { } nodeType)
        {
            // The only allowed node type is element or end of element. All other types should
            // either be skipped (e.g. text) or are errors.
            if (nodeType is XmlNodeType.Element)
            {
                _isStart = true;
                return;
            }

            if (nodeType is XmlNodeType.EndElement)
            {
                _isStart = false;
                return;
            }

            // All other nodes should be skipped:
            // * The possible nodes (Text, Comment, CDATA, SignificantWhitespace,
            //   ProcessingInstruction) should be skipped, because they are not elements. Excel
            //   also skips text that is between nodes where it is not valid, without error.
            // * Other node types are disallowed by usage semantic (Document, None, XmlDeclaration)
            //   or XmlReader setting (DTD).
            // * Attribute should never be encountered because it is after element.
        }
    }

    private XmlNodeType? ReadNode()
    {
        return _reader.Read() ? _reader.NodeType : null;
    }

    private void ThrowWhenReaderNotOnElement()
    {
        if (_reader.NodeType is not XmlNodeType.Element and not XmlNodeType.EndElement)
            throw new InvalidOperationException("XML reader is not on start or end note.");
    }

    private void ThrowOnNonStartElement()
    {
        if (_reader.NodeType != XmlNodeType.Element || !_isStart || _inLookup)
            throw new InvalidOperationException("To read content/attribute, the reader must be on start element and in processing state.");
    }
}
