#nullable disable

using System;
using System.Xml;
using ClosedXML.Extensions;
using static ClosedXML.Excel.XLWorkbook;
using static ClosedXML.Excel.IO.OpenXmlConst;

namespace ClosedXML.Excel.IO
{
    internal class TextSerializer
    {
        internal static void WriteRichTextElements(XmlWriter w, XLImmutableRichText richText, SaveContext context)
        {
            foreach (var textRun in richText.Runs)
            {
                var text = richText.GetRunText(textRun);
                if (text.Length > 0)
                {
                    WriteRun(w, text, textRun.Font);
                }
            }

            if (richText.PhoneticsProperties is not null)
            {
                var phoneticsProps = richText.PhoneticsProperties.Value;
                foreach (var p in richText.PhoneticRuns)
                {
                    w.WriteStartElement("rPh", Main2006SsNs);
                    w.WriteAttribute("sb", p.StartIndex);
                    w.WriteAttribute("eb", p.EndIndex);

                    w.WriteStartElement("t", Main2006SsNs);
                    if (p.Text.PreserveSpaces())
                        w.WritePreserveSpaceAttr();

                    w.WriteString(p.Text);
                    w.WriteEndElement(); // t
                    w.WriteEndElement(); // rPh
                }

                var font = phoneticsProps.Font;
                if (!context.SharedFonts.TryGetValue(font, out FontInfo fi))
                {
                    fi = new FontInfo { Font = font };
                    context.SharedFonts.Add(font, fi);
                }

                w.WriteStartElement("phoneticPr", Main2006SsNs);
                w.WriteAttribute("fontId", fi.FontId);

                if (phoneticsProps.Alignment != XLPhoneticAlignment.Left)
                    w.WriteAttributeString("alignment", phoneticsProps.Alignment.ToOpenXmlString());

                if (phoneticsProps.Type != XLPhoneticType.FullWidthKatakana)
                    w.WriteAttributeString("type", phoneticsProps.Type.ToOpenXmlString());

                w.WriteEndElement(); // phoneticPr
            }
        }

        internal static void WriteRun(XmlWriter w, XLImmutableRichText richText, XLImmutableRichText.RichTextRun run)
        {
            var runText = richText.GetRunText(run);
            WriteRun(w, runText, run.Font);
        }

        private static void WriteRun(XmlWriter w, string text, XLFontValue font)
        {
            w.WriteStartElement("r", Main2006SsNs);
            w.WriteStartElement("rPr", Main2006SsNs);

            if (font.Bold)
                w.WriteEmptyElement("b");

            if (font.Italic)
                w.WriteEmptyElement("i");

            if (font.Strikethrough)
                w.WriteEmptyElement("strike");

            // Three attributes are not stored/written:
            // * outline - doesn't do anything and likely only works in Word.
            // * condense - legacy compatibility setting for macs
            // * extend - legacy compatibility setting for pre-xlsx Excels
            // None have sensible descriptions.

            if (font.Shadow)
                w.WriteEmptyElement("shadow");

            if (font.Underline != XLFontUnderlineValues.None)
                WriteRunProperty(w, "u", font.Underline.ToOpenXmlString());

            WriteRunProperty(w, @"vertAlign", font.VerticalAlignment.ToOpenXmlString());
            WriteRunProperty(w, "sz", font.FontSize);
            w.WriteColor("color", font.FontColor);
            WriteRunProperty(w, "rFont", font.FontName);
            WriteRunProperty(w, "family", (Int32)font.FontFamilyNumbering);

            if (font.FontCharSet != XLFontCharSet.Default)
                WriteRunProperty(w, "charset", (int)font.FontCharSet);

            if (font.FontScheme != XLFontScheme.None)
                WriteRunProperty(w, "scheme", font.FontScheme.ToOpenXml());

            w.WriteEndElement(); // rPr

            w.WriteStartElement("t", Main2006SsNs);
            if (text.PreserveSpaces())
                w.WritePreserveSpaceAttr();

            w.WriteString(text);

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
