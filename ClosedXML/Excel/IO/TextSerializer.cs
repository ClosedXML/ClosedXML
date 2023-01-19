using System;
using System.Linq;
using System.Xml;
using static ClosedXML.Excel.XLWorkbook;
using static ClosedXML.Excel.IO.OpenXmlConst;

namespace ClosedXML.Excel.IO
{
    internal class TextSerializer
    {
        internal static void WriteRichTextElements(XmlWriter w, IXLCell cell, SaveContext context)
        {
            var richText = cell.GetRichText();
            foreach (var rt in richText.Where(r => !String.IsNullOrEmpty(r.Text)))
                WriteRun(w, rt);

            if (richText.HasPhonetics)
            {
                foreach (var p in richText.Phonetics)
                {
                    w.WriteStartElement("rPh", Main2006SsNs);
                    w.WriteAttributeString("sb", p.Start.ToInvariantString());
                    w.WriteAttributeString("eb", p.End.ToInvariantString());

                    w.WriteStartElement("t", Main2006SsNs);
                    if (p.Text.PreserveSpaces())
                    {
                        // TODO: add test
                        w.WriteAttributeString("xml", "space", Xml1998Ns, "preserve");
                    }

                    w.WriteString(p.Text);
                    w.WriteEndElement(); // t
                    w.WriteEndElement(); // rPh
                }

                var fontKey = XLFont.GenerateKey(richText.Phonetics);
                var f = XLFontValue.FromKey(ref fontKey);

                if (!context.SharedFonts.TryGetValue(f, out FontInfo fi))
                {
                    fi = new FontInfo { Font = f };
                    context.SharedFonts.Add(f, fi);
                }

                w.WriteStartElement("phoneticPr", Main2006SsNs);
                w.WriteAttributeString("fontId", fi.FontId.ToInvariantString());

                if (richText.Phonetics.Alignment != XLPhoneticAlignment.Left)
                    w.WriteAttributeString("alignment", richText.Phonetics.Alignment.ToOpenXmlString());

                if (richText.Phonetics.Type != XLPhoneticType.FullWidthKatakana)
                    w.WriteAttributeString("type", richText.Phonetics.Type.ToOpenXmlString());

                w.WriteEndElement(); // phoneticPr
            }
        }

        internal static void WriteRun(XmlWriter w, IXLRichString rt)
        {
            // TODO: Missing outline, charset, condense, extend and scheme properties
            w.WriteStartElement("r", Main2006SsNs);
            w.WriteStartElement("rPr", Main2006SsNs);

            if (rt.Bold)
                WritePropertyElTrue(w, "b");

            if (rt.Italic)
                WritePropertyElTrue(w, "i");

            if (rt.Strikethrough)
                WritePropertyElTrue(w, "strike");

            if (rt.Shadow)
                WritePropertyElTrue(w, "shadow");

            if (rt.Underline != XLFontUnderlineValues.None)
                WriteProperty(w, "u", rt.Underline.ToOpenXmlString());

            WriteProperty(w, "vertAlign", rt.VerticalAlignment.ToOpenXmlString());
            WriteProperty(w, "sz", rt.FontSize);
            WriteColor(w, "color", rt.FontColor);
            WriteProperty(w, "rFont", rt.FontName);
            WriteProperty(w, "family", (Int32)rt.FontFamilyNumbering);

            w.WriteEndElement(); // rPr

            w.WriteStartElement("t", Main2006SsNs);
            if (rt.Text.PreserveSpaces())
            {
                // TODO: add test
                w.WriteAttributeString("xml", "space", Xml1998Ns, "preserve");
            }

            w.WriteString(rt.Text);

            w.WriteEndElement(); // t
            w.WriteEndElement(); // r
        }

        private static void WriteProperty(XmlWriter w, String name, String val)
        {
            w.WriteStartElement(name, Main2006SsNs);
            w.WriteAttributeString("val", val);
            w.WriteEndElement();
        }

        private static void WriteProperty(XmlWriter w, String name, Int32 val)
        {
            w.WriteStartElement(name, Main2006SsNs);
            w.WriteAttributeString("val", val.ToInvariantString());
            w.WriteEndElement();
        }

        private static void WriteProperty(XmlWriter w, String name, Double val)
        {
            w.WriteStartElement(name, Main2006SsNs);
            w.WriteAttributeString("val", val.ToInvariantString());
            w.WriteEndElement();
        }

        private static void WritePropertyElTrue(XmlWriter w, String name)
        {
            w.WriteStartElement(name, Main2006SsNs);
            w.WriteEndElement();
        }

        private static void WriteColor(XmlWriter w, String elName, XLColor xlColor, Boolean isDifferential = false)
        {
            w.WriteStartElement(elName, Main2006SsNs);
            switch (xlColor.ColorType)
            {
                case XLColorType.Color:
                    w.WriteAttributeString("rgb", xlColor.Color.ToHex());
                    break;

                case XLColorType.Indexed:
                    // 64 is 'transparent' and should be ignored for differential formats
                    if (!isDifferential || xlColor.Indexed != 64)
                        w.WriteAttributeString("indexed", xlColor.Indexed.ToInvariantString());
                    break;

                case XLColorType.Theme:
                    w.WriteAttributeString("theme", xlColor.ThemeColor.ToInvariantString());

                    if (xlColor.ThemeTint != 0)
                        w.WriteAttributeString("tint", xlColor.ThemeTint.ToInvariantString());
                    break;
            }

            w.WriteEndElement();
        }
    }
}
