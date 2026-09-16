using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Xml;

namespace ClosedXML.IO;

/// <summary>
/// Markup Compatibility and Extensibility (MCE) processor per ISO-29500-3:2015.
/// </summary>
/// <remarks>
/// Does not process attributes. Fundamentally, it is a facade over <see cref="XmlReader"/> and
/// there is no benefit in skipping attributes. The consuming application asks for a presence of
/// an attribute. If consuming application won't process attribute, it won't even ask fo it.
/// </remarks>
public class MceXmlReader : IXmlReader
{
    private readonly XmlReader _reader;

    // MCE processes every element, so it must be fast. All element/attribute comparisons are done
    // with atomized strings from XmlReader name table.
    private readonly string _mce;
    private readonly XmlName _alternateContent;
    private readonly XmlName _choice;
    private readonly XmlName _fallback;
    private readonly XmlName _attRequires;
    private readonly XmlName _attIgnorable;
    private readonly XmlName _attProcessContent;
    private readonly XmlName _attMustUnderstand;

    private readonly Tracker<string> _ignorable = new(ReferenceEqualityComparer.Instance, static (state, ns) => state.ContainsKey(ns));
    private readonly Tracker<NamePair> _processContent = new(EqualityComparer<NamePair>.Default, static (state, pair) =>
    {
        foreach (var namePair in state.Keys)
        {
            if (namePair.Matches(pair))
                return true;
        }

        return false;
    });

    /// <summary>
    /// A set of namespaces understood by the consumer.
    /// </summary>
    private readonly HashSet<string> _appConfig = new(ReferenceEqualityComparer.Instance);

    /// <summary>
    /// An optional application defined extension element.
    /// </summary>
    private readonly XmlName? _adee;

    /// <summary>
    /// A handler to signal a mismatch.
    /// </summary>
    private readonly Action<MismatchInfo>? _signalMismatch;

    /// <summary>
    /// Is the reader currently in an application defined extension element? The value indicates
    /// depth at which we switched into ADEE mode.
    /// </summary>
    private int? _inAdee;

    private readonly bool _leaveOpen;

    public MceXmlReader(XmlReader reader, MceSettings settings)
        : this(reader, settings, leaveOpen: false)
    {
    }

    public MceXmlReader(XmlReader reader, MceSettings settings, bool leaveOpen)
    {
        const string mceNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";

        _reader = reader;
        var nameTable = _reader.NameTable ?? throw new ArgumentException("XmlReader must use name table.");
        _mce = nameTable.Add(mceNs);
        _alternateContent = XmlName.Atomize("AlternateContent", mceNs, nameTable);
        _choice = XmlName.Atomize("Choice", mceNs, nameTable);
        _fallback = XmlName.Atomize("Fallback", mceNs, nameTable);
        _attRequires = XmlName.Atomize("Requires", "", nameTable);
        _attIgnorable = XmlName.Atomize("Ignorable", mceNs, nameTable);
        _attProcessContent = XmlName.Atomize("ProcessContent", mceNs, nameTable);
        _attMustUnderstand = XmlName.Atomize("MustUnderstand", mceNs, nameTable);

        // Atomize MCE settings
        foreach (var appConfigNs in settings.ApplicationConfiguration)
            _appConfig.Add(nameTable.Add(appConfigNs));

        if (settings.AdeeLocalName is { } extLocalName)
        {
            var atomizedExtName = nameTable.Add(extLocalName);
            var atomizedExtNs = nameTable.Add(settings.AdeeNamespaceUri ?? string.Empty);
            _adee = new XmlName(atomizedExtName, atomizedExtNs);
        }

        _signalMismatch = settings.SignalMismatch;
        _leaveOpen = leaveOpen;

        if (reader.NodeType == XmlNodeType.Element)
        {
            TrackMceAttributes();
            NodeType = XmlTreeNodeType.OpenElement;
        }
    }

    /// <inheritdoc/>
    public XmlTreeNodeType NodeType { get; private set; }

    /// <inheritdoc/>
    public int Depth => _reader.Depth;

    /// <inheritdoc/>
    public string LocalName => _reader.LocalName;

    /// <inheritdoc/>
    public string NamespaceUri => _reader.NamespaceURI;

    /// <inheritdoc/>
    public string Value => _reader.Value;

    /// <inheritdoc/>
    public LineInfo LineInfo => _reader.GetLineInfo();

    /// <inheritdoc/>
    public bool Read()
    {
        if (_inAdee is { } openedAdeeDepth)
        {
            if (!MoveToNextNode())
                throw UnpairedXml();

            if (_reader.Depth == openedAdeeDepth)
                _inAdee = null;

            return true;
        }

        // Loop should end when we are a normal element, not on MCE element. There can be nested AC inside
        while (MoveToNextNode())
        {
            if (_adee is not null && IsOpenElement(_adee.Value))
            {
                _inAdee = _reader.Depth;
                return true;
            }

            if (IsOpenElement(_alternateContent))
            {
                SignalMismatchIfPresent();
                MoveToChoiceOrSkip();
            }
            else if (IsCloseElement(_choice))
            {
                SkipToCloseAlternateContent(false);
            }
            else if (IsCloseElement(_fallback))
            {
                SkipToCloseAlternateContent(true);
            }
            else if (IsIgnored())
            {
                // Ignored and not understood -> skip
                SkipToCloseElement();
            }
            else if (IsUnwrapped())
            {
                // The unwrapped open/close element was consumed and won't be emitted
                if (NodeType == XmlTreeNodeType.OpenElement)
                    SignalMismatchIfPresent();
            }
            else
            {
                // Not part of MCE, return the the consumer
                return true;
            }
        }

        return false;
    }

    /// <inheritdoc/>
    public string? GetAttribute(string attributeName, string? namespaceUri)
    {
        // XmlReader returns to the element node once it reads the attribute value
        return _reader.GetAttribute(attributeName, namespaceUri ?? string.Empty);
    }

    /// <inheritdoc/>
    public void Dispose()
    {
        if (!_leaveOpen)
            _reader.Dispose();
    }

    private bool IsIgnored()
    {
        if (!_ignorable.Declared(_reader.NamespaceURI))
            return false;

        if (_appConfig.Contains(_reader.NamespaceURI))
            return false;

        if (_processContent.Declared(new NamePair(_reader.NamespaceURI, _reader.LocalName)))
            return false;

        return true;
    }

    private bool IsUnwrapped()
    {
        if (!_ignorable.Declared(_reader.NamespaceURI))
            return false;

        if (_appConfig.Contains(_reader.NamespaceURI))
            return false;

        if (!_processContent.Declared(new NamePair(_reader.NamespaceURI, _reader.LocalName)))
            return false;

        return true;
    }

    /// <summary>
    /// Move from <c>AlternateContent</c> open element to selected choice or to the closing
    /// element of <c>AlternateContent</c>. Does not emit nodes (even text ones), because
    /// specification says the <em>replace this AlternateContent element with the content of
    /// the Choice or Fallback element marked as selected.</em>.
    /// </summary>
    private void MoveToChoiceOrSkip()
    {
        Debug.Assert(IsOpenElement(_alternateContent));

        while (MoveToNextNode())
        {
            if (IsCloseElement(_alternateContent))
            {
                return;
            }

            if (IsOpenElement(_choice))
            {
                if (GetAttribute(_attRequires) is not { } requires || string.IsNullOrWhiteSpace(requires))
                    throw MceThrowHelper.InvalidAttribute(_attRequires.LocalName, _reader);

                if (IsChoiceSelected(requires))
                {
                    SignalMismatchIfPresent();
                    return;
                }

                SkipToCloseElement();
            }
            else if (IsOpenElement(_fallback))
            {
                SignalMismatchIfPresent();
                return;
            }
            else if (NodeType == XmlTreeNodeType.OpenElement)
            {
                // AC should only contain only choice/fallback, but to future-proof, it can contain
                // other ignorable elements. Technically, it should also signal mismatch, but it's
                // illegal anyway so it makes no sense to signal mismatch.
                if (!_ignorable.Declared(_reader.NamespaceURI))
                    throw MceThrowHelper.ElementNotIgnorable(_reader.LocalName, _reader);

                SkipToCloseElement();
            }
        }

        throw UnpairedXml();
    }

    private void SkipToCloseAlternateContent(bool seenFallback)
    {
        Debug.Assert(IsCloseElement(_choice) || IsCloseElement(_fallback));
        var depth = _reader.Depth;
        do
        {
            if (!MoveToNextNode())
                throw UnpairedXml();

            if (_reader.Depth == depth && NodeType == XmlTreeNodeType.OpenElement)
            {
                if (IsElement(_fallback))
                {
                    if (seenFallback)
                        throw MceThrowHelper.UnexpectedElementFound(_fallback.LocalName, _reader);

                    seenFallback = true;
                }
                else if (IsElement(_choice))
                {
                    if (seenFallback)
                        throw MceThrowHelper.UnexpectedElementFound(_choice.LocalName, _reader);
                }
                else if (!_ignorable.Declared(_reader.NamespaceURI))
                {
                    throw MceThrowHelper.ElementNotIgnorable(_reader.LocalName, _reader);
                }
            }
        } while (_reader.Depth >= depth);
    }

    private string? GetAttribute(XmlName attributeName)
    {
        if (!_reader.HasAttributes)
            return null;

        // Once done, move back to the element container node, so various checks
        // on XmlReader.NodeType still work.
        while (_reader.MoveToNextAttribute())
        {
            if (ReferenceEquals(_reader.LocalName, attributeName.LocalName) &&
                ReferenceEquals(_reader.NamespaceURI, attributeName.Namespace))
            {
                var value = _reader.Value;
                _reader.MoveToElement();
                return value;
            }
        }

        _reader.MoveToElement();
        return null;
    }

    private bool IsChoiceSelected(string requires)
    {
        var prefixes = requires.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        foreach (var prefix in prefixes)
        {
            var namespaceUri = _reader.LookupNamespace(prefix);
            if (namespaceUri is null)
                throw MceThrowHelper.NamespacePrefixNotFound(_attRequires.LocalName, prefix, _reader);

            if (!_appConfig.Contains(namespaceUri))
                return false;
        }

        return true;
    }

    private bool IsOpenElement(XmlName name)
    {
        return NodeType == XmlTreeNodeType.OpenElement && IsElement(name);
    }

    private bool IsCloseElement(XmlName name)
    {
        return NodeType == XmlTreeNodeType.CloseElement && IsElement(name);
    }

    private bool IsElement(XmlName name)
    {
        Debug.Assert(_reader.NameTable?.Get(name.LocalName) is not null);
        Debug.Assert(_reader.NameTable?.Get(name.Namespace) is not null);
        return ReferenceEquals(_reader.LocalName, name.LocalName) &&
               ReferenceEquals(_reader.NamespaceURI, name.Namespace);
    }

    private void SkipToCloseElement()
    {
        Debug.Assert(NodeType == XmlTreeNodeType.OpenElement);
        var depth = _reader.Depth;
        do
        {
            if (!MoveToNextNode())
                throw UnpairedXml();
        } while (_reader.Depth > depth);
    }

    /// <summary>
    /// Move to next opening or closing element from current element. This is the only permitted method that can move the reader from an element.
    /// </summary>
    private bool MoveToNextNode()
    {
        if (_reader.NodeType == XmlNodeType.Element && _reader.IsEmptyElement && NodeType == XmlTreeNodeType.OpenElement)
        {
            UntrackMceAttributes();
            NodeType = XmlTreeNodeType.CloseElement;
            return true;
        }

        while (_reader.Read())
        {
            switch (_reader.NodeType)
            {
                case XmlNodeType.Element:
                    TrackMceAttributes();
                    NodeType = XmlTreeNodeType.OpenElement;
                    return true;

                case XmlNodeType.EndElement:
                    UntrackMceAttributes();
                    NodeType = XmlTreeNodeType.CloseElement;
                    return true;

                // Text nodes:
                //   The Whitespace node should only appear if XmlReaderSetting.IgnoreWhitespace is
                //   set to false. The 'default' whitespace processing node should depend on the XML
                //   processor, if configuration says give me all whitespaces, give all whitespaces.
                case XmlNodeType.Whitespace:
                case XmlNodeType.Text:
                case XmlNodeType.SignificantWhitespace:
                case XmlNodeType.CDATA:
                case XmlNodeType.EntityReference:
                    NodeType = XmlTreeNodeType.Text;
                    return true;

                // Invalid nodes:
                //   We should never see these. If we do, there is a bug in the reader.
                //   None - If XmlReader.Read() returned false, the reader is never on None node
                //   Attribute - attribute reading code must ensure it ends back on the Element node
                case XmlNodeType.None:
                case XmlNodeType.Attribute:
                    throw new UnreachableException($"Encountered a node {_reader.NodeType}.");

                // Skip nodes:
                //   Nodes that don't produce a XmlTreeNode. Depending on XmlReaderSetting, we might
                //   see them or not, but they don't produce node to consume in any case.
                case XmlNodeType.XmlDeclaration:
                case XmlNodeType.Comment:
                case XmlNodeType.ProcessingInstruction:
                case XmlNodeType.DocumentType:

                // Non-appearing:
                //    The documentation of the XmlReader.NodeType says the property never returns
                //    these types. They only appear in some in-memory XmlDocuments and enum was
                //    used in several places.
                case XmlNodeType.Document:
                case XmlNodeType.DocumentFragment:
                case XmlNodeType.Entity:
                case XmlNodeType.EndEntity:
                case XmlNodeType.Notation:
                default:
                    break;
            }
        }

        NodeType = XmlTreeNodeType.None;
        return false;
    }

    private void TrackMceAttributes()
    {
        Debug.Assert(_reader.NodeType == XmlNodeType.Element);

        if (GetAttribute(_attIgnorable) is { } ignorableValue)
        {
            TrackIgnorable(ignorableValue, _attIgnorable);
        }

        if (GetAttribute(_attProcessContent) is { } processContentValue)
        {
            TrackProcessableContent(processContentValue, _attProcessContent);
        }
    }

    private void TrackIgnorable(string nsList, XmlName attribute)
    {
        foreach (var nsPrefix in nsList.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            if (_reader.LookupNamespace(nsPrefix) is not { } ns)
                throw MceThrowHelper.NamespacePrefixNotFound(attribute.LocalName, nsPrefix, _reader);

            if (ReferenceEquals(ns, _mce))
                throw MceThrowHelper.MceNamespaceNotAllowed(attribute.LocalName, _reader);

            _ignorable.Add(_reader.Depth, ns);
        }
    }

    private void TrackProcessableContent(string processContent, XmlName attribute)
    {
        foreach (var token in processContent.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            var namePair = NamePair.Parse(token, attribute, _reader);
            if (ReferenceEquals(namePair.Namespace, _mce))
                throw MceThrowHelper.MceNamespaceNotAllowed(attribute.LocalName, _reader);

            if (!_ignorable.Declared(namePair.Namespace))
                throw MceThrowHelper.AttributeNamespaceNotIgnorable(attribute.LocalName, namePair.Namespace, _reader);

            _processContent.Add(_reader.Depth, namePair);
        }
    }

    private void UntrackMceAttributes()
    {
        Debug.Assert(_reader.NodeType is XmlNodeType.Element or XmlNodeType.EndElement);
        _ignorable.Clear(_reader.Depth);
        _processContent.Clear(_reader.Depth);
    }

    private void SignalMismatchIfPresent()
    {
        Debug.Assert(NodeType == XmlTreeNodeType.OpenElement);
        if (_signalMismatch is null)
            return;

        if (GetAttribute(_attMustUnderstand) is { } mustUnderstand)
        {
            foreach (var nsPrefix in mustUnderstand.Split(' ', StringSplitOptions.RemoveEmptyEntries))
            {
                if (_reader.LookupNamespace(nsPrefix) is not { } ns)
                    throw MceThrowHelper.NamespacePrefixNotFound(_attMustUnderstand.LocalName, nsPrefix, _reader);

                if (!_appConfig.Contains(ns))
                {
                    var info = new MismatchInfo
                    {
                        LineInfo = _reader.GetLineInfo()
                    };
                    _signalMismatch(info);
                }
            }
        }
    }

    private Exception UnpairedXml()
    {
        // An exception to throw when input is invalid XML (unpaired elements, ends without being
        // at the end of XML tree). That should never happen, because XmlReader should throw when
        // it detects an invalid XML.
        return new UnreachableException($"Not a valid XML stream (unpaired elements) at {_reader.GetLineInfo()}.");
    }

    /// <summary>
    /// A fully qualified XML name of element or an attribute.
    /// </summary>
    /// <param name="LocalName">Local name of an element.</param>
    /// <param name="Namespace">Default namespace is indicated by an empty string.</param>
    private readonly record struct XmlName(string LocalName, string Namespace)
    {
        internal static XmlName Atomize(string localName, string namespaceUri, XmlNameTable nameTable)
        {
            return new XmlName(nameTable.Add(localName), nameTable.Add(namespaceUri));
        }
    };

    /// <summary>
    /// Tracker that keeps track of items encountered on the path from the root to the current
    /// element. It is used to check whether an item was declared on the current element or on
    /// an ancestor element. Because it must determine whether the item matches an item
    /// in the current element or <em>any</em> ancestor, it stores only the item found at
    /// the lowest depth, since items at higher depths are redundant.
    /// </summary>
    private class Tracker<T>
        where T : notnull
    {
        /// <summary>
        /// The key is a item, the value is first depth when it was encountered.
        /// </summary>
        private readonly Dictionary<T, int> _state;
        private readonly HashSet<int> _usedDepths;
        private readonly Func<Dictionary<T, int>, T, bool> _matches;

        internal Tracker(IEqualityComparer<T> comparer, Func<Dictionary<T, int>, T, bool> matches)
        {
            _state = new Dictionary<T, int>(comparer);
            _usedDepths = new HashSet<int>();
            _matches = matches;
        }

        internal void Add(int depth, T item)
        {
            if (_state.TryAdd(item, depth))
                _usedDepths.Add(depth);
        }

        /// <summary>
        /// Was the matching value declared on the current element or an ancestor element?
        /// </summary>
        internal bool Declared(T value)
        {
            return _matches(_state, value);
        }

        /// <summary>
        /// Clear items from current item at depth.
        /// </summary>
        internal void Clear(int depthToClear)
        {
            if (!_usedDepths.Remove(depthToClear))
                return;

            var itemsToRemove = new List<T>();
            foreach (var (item, depth) in _state)
            {
                if (depthToClear == depth)
                    itemsToRemove.Add(item);
            }

            foreach (var item in itemsToRemove)
                _state.Remove(item);
        }
    }

    /// <summary>
    /// Namespace - local name pair for processing content attribute.
    /// </summary>
    private readonly record struct NamePair(string Namespace, string? LocaleName)
    {
        internal static NamePair Parse(string token, XmlName attribute, XmlReader reader)
        {
            var commaIndex = token.IndexOf(':');
            if (commaIndex < 0)
                throw MceThrowHelper.InvalidAttribute(reader.LocalName, reader);

            var nsPrefix = token[..commaIndex];
            if (!IsValidName(nsPrefix))
                throw MceThrowHelper.InvalidAttribute(attribute.LocalName, reader);

            var nameToken = token[(commaIndex + 1)..];
            string? atomizedName;
            if (nameToken is ['*'])
            {
                atomizedName = null;
            }
            else
            {
                if (!IsValidName(nameToken))
                    throw MceThrowHelper.InvalidAttribute(attribute.LocalName, reader);

                atomizedName = reader.NameTable!.Add(nameToken);
            }

            if (reader.LookupNamespace(nsPrefix) is not { } ns)
                throw MceThrowHelper.NamespacePrefixNotFound(attribute.LocalName, nsPrefix, reader);

            return new NamePair(ns, atomizedName);
        }

        private static bool IsValidName(string nameToken)
        {
            return (nameToken.Length > 0 && nameToken.All(XmlConvert.IsNCNameChar));
        }

        internal bool Matches(NamePair other)
        {
            if (ReferenceEquals(Namespace, other.Namespace))
            {
                if (ReferenceEquals(LocaleName, other.LocaleName))
                    return true;

                if (LocaleName is null || other.LocaleName is null)
                    return true;
            }

            return false;
        }
    }
}
