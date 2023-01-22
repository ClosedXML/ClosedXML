using System;
using System.Xml;
using ClosedXML.Extensions;
using static ClosedXML.Excel.XLWorkbook;
using static ClosedXML.Excel.IO.OpenXmlConst;

namespace ClosedXML.Excel.IO
{
    internal class TextSerializer
    {
        internal static void WriteRichTextElements(XmlWriter w, XLCell cell, SaveContext context)
        {
            var richText = cell.GetRichText();
            foreach (var rt in richText)
            {
                if (!String.IsNullOrEmpty(rt.Text))
                {
                    WriteRun(w, rt);
                }
            }

            if (richText.HasPhonetics)
            {
                foreach (var p in richText.Phonetics)
                {
                    w.WriteStartElement("rPh", Main2006SsNs);
                    w.WriteAttribute("sb", p.Start);
                    w.WriteAttribute("eb", p.End);

                    w.WriteStartElement("t", Main2006SsNs);
                    if (p.Text.PreserveSpaces())
                        w.WritePreserveSpaceAttr();

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
                w.WriteAttribute("fontId", fi.FontId);

                if (richText.Phonetics.Alignment != XLPhoneticAlignment.Left)
                    w.WriteAttributeString("alignment", richText.Phonetics.Alignment.ToOpenXmlString());

                if (richText.Phonetics.Type != XLPhoneticType.FullWidthKatakana)
                    w.WriteAttributeString("type", richText.Phonetics.Type.ToOpenXmlString());

                w.WriteEndElement(); // phoneticPr
            }
        }

        internal static void WriteRun(XmlWriter w, XLRichString rt)
        {
            w.WriteStartElement("r", Main2006SsNs);
            w.WriteStartElement("rPr", Main2006SsNs);

            if (rt.Bold)
                w.WriteEmptyElement("b");

            if (rt.Italic)
                w.WriteEmptyElement("i");

            if (rt.Strikethrough)
                w.WriteEmptyElement("strike");

            // Three attributes are not stored/written:
            // * outline - doesn't do anything and likely only works in Word.
            // * condense - legacy compatibility setting for macs
            // * extend - legacy compatibility setting for pre-xlsx Excels
            // None have sensible descriptions.

            if (rt.Shadow)
                w.WriteEmptyElement("shadow");

            if (rt.Underline != XLFontUnderlineValues.None)
                WriteRunProperty(w, "u", rt.Underline.ToOpenXmlString());

            WriteRunProperty(w, @"vertAlign", rt.VerticalAlignment.ToOpenXmlString());
            WriteRunProperty(w, "sz", rt.FontSize);
            w.WriteColor("color", rt.FontColor);
            WriteRunProperty(w, "rFont", rt.FontName);
            WriteRunProperty(w, "family", (Int32)rt.FontFamilyNumbering);

            if (rt.FontCharSet != XLFontCharSet.Default)
                WriteRunProperty(w, "charset", (int)rt.FontCharSet);

            if (rt.FontScheme != XLFontScheme.None)
                WriteRunProperty(w, "scheme", rt.FontScheme.ToOpenXml());

            w.WriteEndElement(); // rPr

            w.WriteStartElement("t", Main2006SsNs);
            if (rt.Text.PreserveSpaces())
                w.WritePreserveSpaceAttr();

            w.WriteString(rt.Text);

            w.WriteEndElement(); // t
            w.WriteEndElement(); // r
        }

        private static void WriteRunProperty(XmlWriter w, String elName, String val)
        {
            w.WriteStartElement(elName, Main2006SsNs);
            w.WriteAttributeString("val", val);
            w.WriteEndElement();
        }

        private static void WriteRunProperty(XmlWriter w, String elName, Int32 val)
        {
            w.WriteStartElement(elName, Main2006SsNs);
            w.WriteAttribute("val", val);
            w.WriteEndElement();
        }

        private static void WriteRunProperty(XmlWriter w, String elName, Double val)
        {
            w.WriteStartElement(elName, Main2006SsNs);
            w.WriteAttribute("val", val);
            w.WriteEndElement();
        }
    }
}
