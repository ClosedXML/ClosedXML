using System;
using System.Diagnostics;
using System.Xml;

namespace ClosedXML.Excel.IO;

/// <summary>
/// <para>
/// Reader that expects that XML document consists of elements in a tree-like fashion. XML
/// shouldn't be mixed, text should be in the leaves (e.g. <c>&lt;f&gt;ABS(A1)*A2&lt;/f&gt;</c>).
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
/// </summary>
public sealed class XmlTreeReader //: IDisposable TODO: Add disposable when Fody is removed.
{
    /// <summary>
    /// The XmlReader that holds current element. The current node should always be either
    /// <see cref="XmlNodeType.Element"/> or <see cref="XmlNodeType.EndElement"/>.
    /// </summary>
    private readonly XmlReader _reader;

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

    // If current element is empty element, this pro has a meaning.
    // if true, it was already opened.
    //private bool emptyIsOpened = false;
    public XmlTreeReader(XmlReader reader)
    {
        _reader = reader;
    }

    /// <summary>
    /// Get name of current element (lookup/processing).
    /// </summary>
    internal string ElementName => _reader.Name;

    /// <summary>
    /// Read next element. Check lookup element is <paramref name="element"/>. If it is, open the
    /// element and return true. Otherwise, return false (element doesn't change).
    /// </summary>
    public bool TryOpen(string element, string ns)
    {
        AssertReaderOnElement();
        SwitchToLookup();

        if (_isStart && _reader.Name == element && _reader.NamespaceURI == ns)
        {
            // Element has been opened, so it should be processed.
            SwitchToProcessing();
            return true;
        }

        return false;
    }

    // Throws when it is on closing elements of incorrect type
    public bool TryClose(string element, string ns)
    {
        AssertReaderOnElement();
        SwitchToLookup();

        if (_isStart || _reader.Name != element || _reader.NamespaceURI != ns)
            return false;

        // Element has been closed, so it should be processed. Though closing elements are not
        // really processed, but we don't want to switch back to lookup. We just want to mark it
        // as "done." Be lazy, e.g. last element of a document doesn't have next element. If
        // parsing logic needs further elements, it will read them when they are needed.
        SwitchToProcessing();
        return true;
    }

    /// <summary>
    /// Assert that we are at the element with <paramref name="elementName"/>. Doesn't move anywhere.
    /// </summary>
    /// <param name="elementName"></param>
    /// <param name="ns"></param>
    public void Open(string elementName, string ns)
    {
        if (!TryOpen(elementName, ns))
            throw PartStructureException.ExpectedElementNotFound($"Expected closing element '{elementName}', but got reader is currently on {(_isStart ? "opening" : "closing")} '{_reader.Name}'.");
    }

    /// <summary>
    /// Close the next unprocessed node. If the node doesn't match the <paramref name="element"/>,
    /// throw an exception.
    /// </summary>
    public void Close(string element, string ns)
    {
        if (!TryClose(element, ns))
            throw PartStructureException.ExpectedElementNotFound();
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

        if (_reader.IsEmptyElement && _isStart)
        {
            // Keep element, but interpret it as the ending element.
            _inLookup = true;
            _isStart = false;
            return;
        }

        // Read next element.
        _reader.Read();
        _reader.MoveToContent();

        _inLookup = true;
        _isStart = _reader.NodeType switch
        {
            XmlNodeType.Element => true,
            XmlNodeType.EndElement => false,
            _ => throw PartStructureException.ExpectedElementNotFound($"Parser expected an element, instead found node '{_reader.NodeType}'."),
        };
    }

    private void AssertReaderOnElement()
    {
        // Use Debug.Assert, so the release version eliminates whole call.
        Debug.Assert(_reader.NodeType is XmlNodeType.Element or XmlNodeType.EndElement);
    }

    // Reader should be on opening node of an element. Skip to the closing and after
    public void Skip()
    {
        AssertReaderOnElement();
        Debug.Assert(!_inLookup);

        // Skip everything under current element, including end element.
        _reader.Skip();

        // We have likely ended up on whitespace, move to element
        SwitchToLookup();
    }

    public bool? GetBool(string attributeName)
    {
        Debug.Assert(!_inLookup);

        _reader.MoveToAttribute(attributeName);
        var result = _reader.ReadContentAsBoolean();
        _reader.MoveToElement();
        return result;
    }

    public bool? GetOptionalBool(string attributeName)
    {
        Debug.Assert(!_inLookup);
        bool? result = _reader.MoveToAttribute(attributeName) ? _reader.ReadContentAsBoolean() : null;
        _reader.MoveToElement();
        return result;
    }

    public int GetInt(string attributeName)
    {
        Debug.Assert(!_inLookup);
        _reader.MoveToAttribute(attributeName);
        var number = _reader.ReadContentAsInt();
        _reader.MoveToElement();
        return number;
    }

    public int? GetOptionalUint(string attributeName)
    {
        Debug.Assert(!_inLookup);
        int? number = _reader.MoveToAttribute(attributeName) ? _reader.ReadContentAsInt() : null;
        if (number < 0)
            throw PartStructureException.InvalidAttributeValue(_reader.ReadContentAsString());

        _reader.MoveToElement();
        return number;
    }

    public int GetUint(string attributeName)
    {
        var value = GetOptionalUint(attributeName);
        if (value is null)
            throw PartStructureException.RequiredElementIsMissing(attributeName);

        return value.Value;
    }

    public double? GetOptionalDouble(string attributeName, double? defaultValue)
    {
        Debug.Assert(!_inLookup);
        var number = _reader.MoveToAttribute(attributeName) ? _reader.ReadContentAsDouble() : defaultValue;
        _reader.MoveToElement();
        return number;
    }
    public double GetDouble(string attributeName)
    {
        Debug.Assert(!_inLookup);
        _reader.MoveToAttribute(attributeName);
        var number = _reader.ReadContentAsDouble();
        _reader.MoveToElement();
        return number;
    }

    public string GetAsXString(string attributeName)
    {
        // TODO: Decode XString
        Debug.Assert(!_inLookup);
        if (!_reader.MoveToAttribute(attributeName))
            throw PartStructureException.RequiredAttributeIsMissing(attributeName, _reader);

        var text = _reader.ReadContentAsString();
        _reader.MoveToElement();
        return text;
    }

    public string? GetOptionalString(string attributeName)
    {
        Debug.Assert(!_inLookup);
        return _reader.GetAttribute(attributeName);
    }

    public TEnum GetEnum<TEnum>(string attributeName)
        where TEnum : struct, Enum
    {
        Debug.Assert(!_inLookup);
        if (!_reader.MoveToAttribute(attributeName))
            throw PartStructureException.RequiredAttributeIsMissing(attributeName, _reader);

        var enumString = _reader.ReadContentAsString();
        _reader.MoveToElement();

        if (!XmlToEnumMapper.TryGetEnum<TEnum>(enumString, out var enumValue))
            throw PartStructureException.InvalidAttributeValue(enumString);

        return enumValue;
    }

    public TEnum GetOptionalEnum<TEnum>(string attributeName, TEnum defaultValue)
        where TEnum : struct, Enum
    {
        Debug.Assert(!_inLookup);
        var enumString = _reader.MoveToAttribute(attributeName) ? _reader.ReadContentAsString() : null;
        _reader.MoveToElement();

        if (enumString is null)
            return defaultValue;

        if (!XmlToEnumMapper.TryGetEnum<TEnum>(enumString, out var enumValue))
            throw PartStructureException.InvalidAttributeValue(enumString);

        return enumValue;
    }
}
