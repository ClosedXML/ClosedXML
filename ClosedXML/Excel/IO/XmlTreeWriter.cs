using ClosedXML.Extensions;
using ClosedXML.IO;
using System;
using System.Xml;

namespace ClosedXML.Excel.IO;

[Janitor.SkipWeaving]
internal class XmlTreeWriter : IDisposable
{
    private readonly XmlWriter _xml;
    private readonly IEnumMapper _enumMapper;

    internal XmlTreeWriter(XmlWriter xml, IEnumMapper enumMapper)
    {
        _xml = xml;
        _enumMapper = enumMapper;
    }

    public void WriteStartDocument(string rootElementName, string ns)
    {
        _xml.WriteStartDocument();

        // Make root element ns a default namespace to avoid prefix if possible
        _xml.WriteStartElement(rootElementName, ns);
        _xml.WriteAttributeString("xmlns", ns);
    }

    public void WriteStartElement(string localName, string ns)
    {
        _xml.WriteStartElement(localName, ns);
    }

    public void WriteAttribute(string attributeName, int value)
    {
        _xml.WriteAttribute(attributeName, value);
    }

    public void WriteAttribute(string attributeName, bool value)
    {
        _xml.WriteAttribute(attributeName, value);
    }

    public void WriteAttribute(string attributeName, string value)
    {
        _xml.WriteAttribute(attributeName, value);
    }

    public void WriteAttribute(string attributeName, double value)
    {
        _xml.WriteAttribute(attributeName, value);
    }

    public void WriteAttribute<TEnum>(string attributeName, TEnum value)
        where TEnum : struct, Enum
    {
        if (!_enumMapper.TryGetText(value, out var text))
            throw new InvalidOperationException($"Missing mapping for enum {value} ({typeof(TEnum).Name}).");

        _xml.WriteAttribute(attributeName, text);
    }

    public void WriteEndElement()
    {
        _xml.WriteEndElement();
    }

    public void Dispose()
    {
        _xml.Dispose();
    }
}
