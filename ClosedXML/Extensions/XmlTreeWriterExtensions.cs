using System;
using ClosedXML.Excel;
using ClosedXML.Excel.IO;

namespace ClosedXML.Extensions;

/// <summary>
/// Extension methods for <see cref="XmlTreeWriter"/>. Keep the original class reasonably clean
/// and place convenience method here.
/// </summary>
internal static class XmlTreeWriterExtensions
{
    /// <summary>
    /// Write a <c>CT_BooleanProperty</c> element.
    /// </summary>
    public static void WriteBooleanProperty(this XmlTreeWriter xml, string elName, bool value, string ns)
    {
        xml.WriteStartElement(elName, ns);
        xml.WriteAttributeDefault("val", value, true);
        xml.WriteEndElement();
    }

    public static void WriteAttributeDefault(this XmlTreeWriter xml, string attrName, bool value, bool defaultValue)
    {
        if (value != defaultValue)
            xml.WriteAttribute(attrName, value);
    }

    public static void WriteAttributeDefault(this XmlTreeWriter xml, string attrName, int value, int defaultValue)
    {
        if (value != defaultValue)
            xml.WriteAttribute(attrName, value);
    }

    public static void WriteAttributeDefault(this XmlTreeWriter xml, string attrName, uint value, uint defaultValue)
    {
        if (value != defaultValue)
            xml.WriteAttribute(attrName, value);
    }

    public static void WriteAttributeDefault(this XmlTreeWriter xml, string attrName, double value, double defaultValue)
    {
        if (!value.Equals(defaultValue))
            xml.WriteAttribute(attrName, value);
    }

    public static void WriteAttributeDefault<TEnum>(this XmlTreeWriter xml, string attrName, TEnum enumValue, TEnum defaultValue)
        where TEnum : struct, Enum
    {
        if (!enumValue.Equals(defaultValue))
            xml.WriteAttribute(attrName, enumValue);
    }

    public static void WriteAttributeOptional(this XmlTreeWriter xml, string attrName, int? value)
    {
        if (value is not null)
            xml.WriteAttribute(attrName, value.Value);
    }

    public static void WriteColor(this XmlTreeWriter xml, string elName, string ns, XLColor color, bool isDifferential = false)
    {
        xml.WriteStartElement(elName, ns);
        switch (color.ColorType)
        {
            case XLColorType.Color:
                xml.WriteAttribute("rgb", color.Color.ToHex());
                break;

            case XLColorType.Indexed:
                // 64 is 'transparent' and should be ignored for differential formats
                if (!isDifferential || color.Indexed != 64)
                    xml.WriteAttribute("indexed", color.Indexed);
                break;

            case XLColorType.Theme:
                xml.WriteAttribute("theme", (int)color.ThemeColor);

                if (color.ThemeTint != 0)
                    xml.WriteAttribute("tint", color.ThemeTint);
                break;
        }

        xml.WriteEndElement();
    }
}
