using System;
using System.Xml;
using ClosedXML.Excel;
using ClosedXML.Excel.IO;

namespace ClosedXML.Extensions
{
    internal static class XmlWriterExtensions
    {
        public static void WriteAttribute(this XmlWriter w, String attrName, String value)
        {
            w.WriteStartAttribute(attrName);
            w.WriteValue(value);
            w.WriteEndAttribute();
        }

        public static void WriteAttributeOptional(this XmlWriter w, String attrName, String? value)
        {
            if (!string.IsNullOrEmpty(value))
                w.WriteAttribute(attrName, value);
        }

        public static void WriteAttribute(this XmlWriter w, String attrName, Int32 value)
        {
            w.WriteStartAttribute(attrName);
            w.WriteValue(value);
            w.WriteEndAttribute();
        }

        public static void WriteAttribute(this XmlWriter w, String attrName, UInt32 value)
        {
            w.WriteStartAttribute(attrName);
            w.WriteValue(value);
            w.WriteEndAttribute();
        }

        public static void WriteAttributeOptional(this XmlWriter w, String attrName, UInt32? value)
        {
            if (value is not null)
                w.WriteAttribute(attrName, value.Value);
        }

        public static void WriteAttributeOptional(this XmlWriter w, String attrName, Int32? value)
        {
            if (value is not null)
                w.WriteAttribute(attrName, value.Value);
        }

        public static void WriteAttribute(this XmlWriter w, String attrName, Double value)
        {
            w.WriteStartAttribute(attrName);
            w.WriteNumberValue(value);
            w.WriteEndAttribute();
        }

        public static void WriteAttribute(this XmlWriter w, String attrName, Boolean value)
        {
            w.WriteStartAttribute(attrName);
            w.WriteValue(value ? "1" : "0");
            w.WriteEndAttribute();
        }

        public static void WriteAttributeDefault(this XmlWriter w, String attrName, Boolean value, Boolean defaultValue)
        {
            if (value != defaultValue)
                w.WriteAttribute(attrName, value);
        }

        public static void WriteAttributeOptional(this XmlWriter w, String attrName, Boolean? value)
        {
            if (value is not null)
                w.WriteAttribute(attrName, value.Value);
        }

        public static void WriteAttributeDefault(this XmlWriter w, String attrName, int value, int defaultValue)
        {
            if (value != defaultValue)
                w.WriteAttribute(attrName, value);
        }

        public static void WriteAttributeDefault(this XmlWriter w, String attrName, uint value, uint defaultValue)
        {
            if (value != defaultValue)
                w.WriteAttribute(attrName, value);
        }

        /// <summary>
        /// Write date in a format <c>2015-01-01T00:00:00</c> (ignore kind).
        /// </summary>
        public static void WriteAttribute(this XmlWriter w, String attrName, DateTime value)
        {
            w.WriteStartAttribute(attrName);
            w.WriteValue(value.ToString("s"));
            w.WriteEndAttribute();
        }

        public static void WriteAttribute(this XmlWriter w, String attrName, String ns, Double value)
        {
            w.WriteStartAttribute(attrName, ns);
            w.WriteNumberValue(value);
            w.WriteEndAttribute();
        }

        public static void WriteNumberValue(this XmlWriter w, Double value)
        {
            // G17 will survive roundtrip to file and back
            w.WriteValue(value.ToInvariantString());
        }

        public static void WritePreserveSpaceAttr(this XmlWriter w)
        {
            w.WriteAttributeString("xml", "space", OpenXmlConst.Xml1998Ns, "preserve");
        }

        public static void WriteEmptyElement(this XmlWriter w, String elName)
        {
            w.WriteStartElement(elName, OpenXmlConst.Main2006SsNs);
            w.WriteEndElement();
        }

        public static void WriteColor(this XmlWriter w, String elName, XLColor xlColor, Boolean isDifferential = false)
        {
            w.WriteStartElement(elName, OpenXmlConst.Main2006SsNs);
            switch (xlColor.ColorType)
            {
                case XLColorType.Color:
                    w.WriteAttributeString("rgb", xlColor.Color.ToHex());
                    break;

                case XLColorType.Indexed:
                    // 64 is 'transparent' and should be ignored for differential formats
                    if (!isDifferential || xlColor.Indexed != 64)
                        w.WriteAttribute("indexed", xlColor.Indexed);
                    break;

                case XLColorType.Theme:
                    w.WriteAttribute("theme", (int)xlColor.ThemeColor);

                    if (xlColor.ThemeTint != 0)
                        w.WriteAttribute("tint", xlColor.ThemeTint);
                    break;
            }

            w.WriteEndElement();
        }
    }
}
