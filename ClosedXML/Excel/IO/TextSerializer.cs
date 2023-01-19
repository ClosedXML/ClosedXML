using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System;
using System.Linq;
using System.Xml;
using ClosedXML.Utils;
using static ClosedXML.Excel.XLWorkbook;
using static ClosedXML.Excel.IO.OpenXmlConst;

namespace ClosedXML.Excel.IO
{
    internal class TextSerializer
    {
        internal static void PopulatedRichTextElements(XmlWriter w, RstType rstType, IXLCell cell, SaveContext context)
        {
            var richText = cell.GetRichText();
            foreach (var rt in richText.Where(r => !String.IsNullOrEmpty(r.Text)))
            {
                var run = GetRun(w, rt);
                rstType.Append(run);
            }

            if (richText.HasPhonetics)
            {
                foreach (var p in richText.Phonetics)
                {
                    var phoneticRun = new PhoneticRun
                    {
                        BaseTextStartIndex = (UInt32)p.Start,
                        EndingBaseIndex = (UInt32)p.End
                    };

                    w.WriteStartElement("rPh", Main2006SsNs);
                    w.WriteAttributeString("sb", p.Start.ToInvariantString());
                    w.WriteAttributeString("eb", p.End.ToInvariantString());

                    w.WriteStartElement("t", Main2006SsNs);

                    var text = new Text { Text = p.Text };
                    if (p.Text.PreserveSpaces())
                    {
                        // TODO: add test
                        w.WriteAttributeString("xml", "space", Xml1998Ns, "preserve");
                        text.Space = SpaceProcessingModeValues.Preserve;
                    }

                    w.WriteString(p.Text);
                    w.WriteEndElement(); // t
                    w.WriteEndElement(); // rPh

                    phoneticRun.Append(text);
                    rstType.Append(phoneticRun);
                }

                var fontKey = XLFont.GenerateKey(richText.Phonetics);
                var f = XLFontValue.FromKey(ref fontKey);

                if (!context.SharedFonts.TryGetValue(f, out FontInfo fi))
                {
                    fi = new FontInfo { Font = f };
                    context.SharedFonts.Add(f, fi);
                }

                var phoneticProperties = new PhoneticProperties
                {
                    FontId = fi.FontId
                };

                w.WriteStartElement("phoneticPr", Main2006SsNs);
                w.WriteAttributeString("fontId", fi.FontId.ToInvariantString());

                if (richText.Phonetics.Alignment != XLPhoneticAlignment.Left)
                {
                    w.WriteAttributeString("alignment", richText.Phonetics.Alignment.ToOpenXmlString());
                    phoneticProperties.Alignment = richText.Phonetics.Alignment.ToOpenXml();
                }

                if (richText.Phonetics.Type != XLPhoneticType.FullWidthKatakana)
                {
                    w.WriteAttributeString("type", richText.Phonetics.Type.ToOpenXmlString());
                    phoneticProperties.Type = richText.Phonetics.Type.ToOpenXml();
                }

                w.WriteEndElement(); // phoneticPr

                rstType.Append(phoneticProperties);
            }
        }

        internal static Run GetRun(XmlWriter w, IXLRichString rt)
        {
            // TODO: Missing outline, charset, condense, extend and scheme properties
            var run = new Run();
            w.WriteStartElement("r", Main2006SsNs);

            var runProperties = new RunProperties();
            w.WriteStartElement("rPr", Main2006SsNs);

            var bold = rt.Bold ? new Bold() : null;
            if (bold != null)
            {
                runProperties.Append(bold);
                WritePropertyElTrue(w, "b");
            }

            var italic = rt.Italic ? new Italic() : null;
            if (italic != null)
            {
                runProperties.Append(italic);
                WritePropertyElTrue(w, "i");
            }

            var strike = rt.Strikethrough ? new Strike() : null;
            if (strike != null)
            {
                runProperties.Append(strike);
                WritePropertyElTrue(w, "strike");
            }

            var shadow = rt.Shadow ? new Shadow() : null;
            if (shadow != null)
            {
                runProperties.Append(shadow);
                WritePropertyElTrue(w, "shadow");
            }

            var underline = rt.Underline != XLFontUnderlineValues.None
                ? new Underline { Val = rt.Underline.ToOpenXml() }
                : null;
            if (underline != null)
            {
                runProperties.Append(underline);
                WriteProperty(w, "u", rt.Underline.ToOpenXmlString());
            }

            var verticalAlignment = new VerticalTextAlignment
            { Val = rt.VerticalAlignment.ToOpenXml() };
            runProperties.Append(verticalAlignment);
            WriteProperty(w, "vertAlign", rt.VerticalAlignment.ToOpenXmlString());

            var fontSize = new FontSize { Val = rt.FontSize };
            runProperties.Append(fontSize);
            WriteProperty(w, "sz", rt.FontSize);

            var color = new Color().FromClosedXMLColor<Color>(rt.FontColor);
            runProperties.Append(color);
            WriteColor(w, "color", rt.FontColor);

            var fontName = new RunFont { Val = rt.FontName };
            runProperties.Append(fontName);
            WriteProperty(w, "rFont", rt.FontName);

            var fontFamilyNumbering = new FontFamily { Val = (Int32)rt.FontFamilyNumbering };
            runProperties.Append(fontFamilyNumbering);
            WriteProperty(w, "family", (Int32)rt.FontFamilyNumbering);

            w.WriteEndElement(); // rPr

            var text = new Text { Text = rt.Text };
            w.WriteStartElement("t", Main2006SsNs);
            if (rt.Text.PreserveSpaces())
            {
                text.Space = SpaceProcessingModeValues.Preserve;

                // TODO: add test
                w.WriteAttributeString("xml", "space", Xml1998Ns, "preserve");
            }

            w.WriteString(rt.Text);

            w.WriteEndElement(); // t
            w.WriteEndElement(); // r
            run.Append(runProperties);
            run.Append(text);
            return run;
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
