using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System;
using System.Linq;
using ClosedXML.Utils;
using static ClosedXML.Excel.XLWorkbook;

namespace ClosedXML.Excel.IO
{
    internal class TextSerializer
    {
        internal static void PopulatedRichTextElements(RstType rstType, IXLCell cell, SaveContext context)
        {
            var richText = cell.GetRichText();
            foreach (var rt in richText.Where(r => !String.IsNullOrEmpty(r.Text)))
            {
                rstType.Append(GetRun(rt));
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

                    var text = new Text { Text = p.Text };
                    if (p.Text.PreserveSpaces())
                        text.Space = SpaceProcessingModeValues.Preserve;

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

                if (richText.Phonetics.Alignment != XLPhoneticAlignment.Left)
                    phoneticProperties.Alignment = richText.Phonetics.Alignment.ToOpenXml();

                if (richText.Phonetics.Type != XLPhoneticType.FullWidthKatakana)
                    phoneticProperties.Type = richText.Phonetics.Type.ToOpenXml();

                rstType.Append(phoneticProperties);
            }
        }
        
        internal static Run GetRun(IXLRichString rt)
        {
            var run = new Run();

            var runProperties = new RunProperties();

            var bold = rt.Bold ? new Bold() : null;
            var italic = rt.Italic ? new Italic() : null;
            var underline = rt.Underline != XLFontUnderlineValues.None
                ? new Underline { Val = rt.Underline.ToOpenXml() }
                : null;
            var strike = rt.Strikethrough ? new Strike() : null;
            var verticalAlignment = new VerticalTextAlignment
                { Val = rt.VerticalAlignment.ToOpenXml() };
            var shadow = rt.Shadow ? new Shadow() : null;
            var fontSize = new FontSize { Val = rt.FontSize };
            var color = new Color().FromClosedXMLColor<Color>(rt.FontColor);
            var fontName = new RunFont { Val = rt.FontName };
            var fontFamilyNumbering = new FontFamily { Val = (Int32)rt.FontFamilyNumbering };

            if (bold != null) runProperties.Append(bold);
            if (italic != null) runProperties.Append(italic);

            if (strike != null) runProperties.Append(strike);
            if (shadow != null) runProperties.Append(shadow);
            if (underline != null) runProperties.Append(underline);
            runProperties.Append(verticalAlignment);

            runProperties.Append(fontSize);
            runProperties.Append(color);
            runProperties.Append(fontName);
            runProperties.Append(fontFamilyNumbering);

            var text = new Text { Text = rt.Text };
            if (rt.Text.PreserveSpaces())
                text.Space = SpaceProcessingModeValues.Preserve;

            run.Append(runProperties);
            run.Append(text);
            return run;
        }
    }
}
