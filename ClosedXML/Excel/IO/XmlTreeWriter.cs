using System;
using System.Xml;
using ClosedXML.Extensions;
using ClosedXML.IO;

namespace ClosedXML.Excel.IO;

internal sealed class XmlTreeWriter : IDisposable
{
    private readonly XmlWriter _xml;
    private readonly IEnumMapper _enumMapper;
    private bool _disposed;

    internal XmlTreeWriter(XmlWriter xml, IEnumMapper enumMapper)
    {
        _xml = xml;
        _enumMapper = enumMapper;
    }

    public void WriteStartDocument(string rootElementName, string ns)
    {
        ThrowIfDisposed();

        // No part should rely on external DTD, plus Excel also writes standalone="yes"
        _xml.WriteStartDocument(standalone: true);

        // Make root element ns a default namespace to avoid prefix if possible
        _xml.WriteStartElement(rootElementName, ns);
        _xml.WriteAttributeString("xmlns", ns);
    }

    public void WriteStartElement(string localName, string ns)
    {
        ThrowIfDisposed();
        _xml.WriteStartElement(localName, ns);
    }

    public void WriteStartExtension(string extUri, string defaultNs, string nsPrefix, string extNs)
    {
        ThrowIfDisposed();
        WriteStartElement("ext", defaultNs);
        WriteAttribute("uri", extUri);
        WriteNsPrefix(nsPrefix, extNs);
    }

    public void WriteNsPrefix(string prefix, string ns)
    {
        ThrowIfDisposed();
        _xml.WriteAttributeString("xmlns", prefix, null, ns);
    }

    public void WriteAttribute(string attributeName, int value)
    {
        ThrowIfDisposed();
        _xml.WriteAttribute(attributeName, value);
    }

    public void WriteAttribute(string attributeName, bool value)
    {
        ThrowIfDisposed();
        _xml.WriteAttribute(attributeName, value);
    }

    public void WriteAttribute(string attributeName, string value)
    {
        ThrowIfDisposed();
        _xml.WriteAttribute(attributeName, value);
    }

    public void WriteAttribute(string attributeName, double value)
    {
        ThrowIfDisposed();
        _xml.WriteAttribute(attributeName, value);
    }

    public void WriteAttribute<TEnum>(string attributeName, TEnum value)
        where TEnum : struct, Enum
    {
        ThrowIfDisposed();
        if (!_enumMapper.TryGetText(value, out var text))
            throw new InvalidOperationException($"Missing mapping for enum {value} ({typeof(TEnum).Name}).");

        _xml.WriteAttribute(attributeName, text);
    }

    public void WriteEndElement()
    {
        ThrowIfDisposed();
        _xml.WriteEndElement();
    }

    public void WriteEndDocument()
    {
        ThrowIfDisposed();
        _xml.WriteEndElement();
        _xml.WriteEndDocument();
    }

    public void Dispose()
    {
        _xml.Dispose();
        _disposed = true;
    }

    private void ThrowIfDisposed()
    {
        if (_disposed)
            throw new ObjectDisposedException(nameof(XmlTreeWriter));
    }
}
