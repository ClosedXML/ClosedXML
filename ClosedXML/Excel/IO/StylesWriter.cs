#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using ClosedXML.Excel.Formatting;
using ClosedXML.Extensions;
using ClosedXML.IO;
using ClosedXML.Utils;
using DocumentFormat.OpenXml.Packaging;
using static ClosedXML.Excel.IO.OpenXmlConst;

namespace ClosedXML.Excel.IO;

internal class StylesWriter
{
    private const int FirstUserDefinedFormatIndex = 164;
    private readonly string _ns = Main2006SsNs;

    internal void WriteContent(WorkbookStylesPart stylesPart, IEnumMapper mapper, XLWorkbookStyles styles, XLWorkbook.SaveContext context)
    {
        // Determine which format components are used and thus should be saved.
        // TODO: For now just assume everything in styles is used
        var usedCellFormats = styles.CellFormats.Select(x => x.Value).ToHashSet();
        var usedNumberFormats = new HashSet<string>();
        var usedFonts = new HashSet<XLFontFormatValue>();
        var usedFills = new HashSet<XLFillFormatValue>();
        var usedBorders = new HashSet<XLBorderFormatValue>();
        foreach (var cellFormat in usedCellFormats)
        {
            if (cellFormat.NumberFormat is { } numberFormat)
                usedNumberFormats.Add(numberFormat);

            if (cellFormat.Font is { } font)
                usedFonts.Add(font);

            if (cellFormat.Fill is { } fill)
                usedFills.Add(fill);

            if (cellFormat.Border is { } border)
                usedBorders.Add(border);
        }

        var settings = new XmlWriterSettings
        {
            Encoding = XLHelper.NoBomUTF8
        };

        using var partStream = stylesPart.GetStream(FileMode.Create);
        using var xml = new XmlTreeWriter(XmlWriter.Create(partStream, settings), mapper);

        xml.WriteStartDocument("styleSheet", _ns);

        // Number formats
        var numberFormatMap = new SequentialMap<int, string>(styles.NumberFormats);

        // Add identity map for predefined formats. That way they can be always be mapped. If there
        // is a workbook that uses predefined ids for a different format (e.g., numId=0 with format
        // "0.00"), it should be dealt in the loading code.
        for (var i = 0; i < FirstUserDefinedFormatIndex; ++i)
            numberFormatMap.Add(i);

        foreach (var (actualId, format) in styles.NumberFormats)
        {
            if (!usedNumberFormats.Contains(format))
                continue;

            // Predefined formats were already added in the identity map and don't need to be written.
            if (XLPredefinedFormat.NumberFormatIds.ContainsKey(format))
                continue;

            numberFormatMap.Add(actualId);
        }

        numberFormatMap.Sort();
        if (numberFormatMap.Count > FirstUserDefinedFormatIndex)
            WriteNumberFormats(xml, numberFormatMap);

        // Fonts
        var fontFormatsMap = SequentialMap<int, XLFontFormatValue>.Create(usedFonts, styles.Fonts);
        if (fontFormatsMap.Count > 0)
            WriteFonts(xml, fontFormatsMap);

        // TODO: Add the fixed fills (none+gray125) at the start
        var fillsFormatsMap = SequentialMap<int, XLFillFormatValue>.Create(usedFills, styles.Fills);
        if (fillsFormatsMap.Count > 0)
            WriteFills(xml, fillsFormatsMap);

        var borderFormatsMap = SequentialMap<int, XLBorderFormatValue>.Create(usedBorders, styles.Borders);
        if (borderFormatsMap.Count > 0)
            WriteBorders(xml, borderFormatsMap);

        // TODO: Ensure the normal style cellXfs has index 0
        var cellXfsMap = SequentialMap<int, XLCellFormatValue>.Create(usedCellFormats, styles.CellFormats);
        if (cellXfsMap.Count > 0)
            WriteCellXfs(xml, cellXfsMap, numberFormatMap, fontFormatsMap, fillsFormatsMap, borderFormatsMap);

        // TODO: Rest of elements cellStyleXfs, cellStyles, dxfs, tableStyles, colors
        xml.WriteEndElement();
    }

    private void WriteNumberFormats(XmlTreeWriter xml, SequentialMap<int, string> idMap)
    {
        xml.WriteStartElement("numFmts", _ns);
        xml.WriteAttribute("count", idMap.Count - FirstUserDefinedFormatIndex);

        foreach (var (savedId, format) in idMap.GetActual())
        {
            // The number format map has identity map for predefined formats
            if (savedId < FirstUserDefinedFormatIndex)
                continue;

            xml.WriteStartElement("numFmt", _ns);
            xml.WriteAttribute("numFmtId", savedId);
            xml.WriteAttribute("formatCode", format);
            xml.WriteEndElement();
        }

        xml.WriteEndElement(); // numFmts
    }

    private void WriteFonts(XmlTreeWriter xml, SequentialMap<int, XLFontFormatValue> idMap)
    {
        xml.WriteStartElement("fonts", _ns);
        xml.WriteAttribute("count", idMap.Count);

        foreach (var (_, font) in idMap.GetActual())
        {
            // MS-OI29500 dictates font elements order.
            xml.WriteStartElement("font", _ns);

            if (font.Bold is { } bold)
                xml.WriteBooleanProperty("b", bold, _ns);

            if (font.Italic is { } italic)
                xml.WriteBooleanProperty("i", italic, _ns);

            if (font.Strikethrough is { } strikethrough)
                xml.WriteBooleanProperty("strike", strikethrough, _ns);

            if (font.Condense is { } condense)
                xml.WriteBooleanProperty("condense", condense, _ns);

            if (font.Extend is { } extend)
                xml.WriteBooleanProperty("extend", extend, _ns);

            if (font.Outline is { } outline)
                xml.WriteBooleanProperty("outline", outline, _ns);

            if (font.Shadow is { } shadow)
                xml.WriteBooleanProperty("shadow", shadow, _ns);

            if (font.Underline is { } underline)
            {
                xml.WriteStartElement("u", _ns);
                xml.WriteAttributeDefault("val", underline, XLFontUnderlineValues.Single);
                xml.WriteEndElement();
            }

            if (font.VerticalAlignment is { } verticalAlignment)
            {
                xml.WriteStartElement("vertAlign", _ns);
                xml.WriteAttribute("val", verticalAlignment);
                xml.WriteEndElement();
            }

            if (font.Size is { } size)
            {
                xml.WriteStartElement("sz", _ns);
                xml.WriteAttribute("val", size.Points);
                xml.WriteEndElement();
            }

            if (font.Color is { } color)
            {
                xml.WriteColor("color", _ns, color);
            }

            if (font.Name is { } name)
            {
                xml.WriteStartElement("name", _ns);
                xml.WriteAttribute("val", name.Text);
                xml.WriteEndElement();
            }

            if (font.Family is { } family)
            {
                xml.WriteStartElement("family", _ns);
                xml.WriteAttribute("val", (int)family);
                xml.WriteEndElement();
            }

            if (font.Charset is { } charset)
            {
                // Charset is stored as an CT_IntProperty
                xml.WriteStartElement("charset", _ns);
                xml.WriteAttribute("val", (int)charset);
                xml.WriteEndElement();
            }

            if (font.Scheme is { } scheme)
            {
                xml.WriteStartElement("scheme", _ns);
                xml.WriteAttribute("val", scheme);
                xml.WriteEndElement();
            }

            xml.WriteEndElement();
        }

        xml.WriteEndElement(); // fonts
    }

    private void WriteFills(XmlTreeWriter xml, SequentialMap<int, XLFillFormatValue> idMap)
    {
        xml.WriteStartElement("fills", _ns);
        xml.WriteAttribute("count", idMap.Count);

        foreach (var (_, fill) in idMap.GetActual())
        {
            xml.WriteStartElement("fill", _ns);
            if (fill.Pattern is { } patternFill)
            {
                xml.WriteStartElement("patternFill", _ns);
                xml.WriteAttribute("patternType", patternFill.PatternType);
                xml.WriteColor("fgColor", _ns, patternFill.PatternColor);
                xml.WriteColor("bgColor", _ns, patternFill.BackgroundColor);
                xml.WriteEndElement();
            }
            else if (fill.LinearGradient is { } linearGradient)
            {
                throw new NotImplementedException();
            }
            else if (fill.PathGradient is { } pathGradient)
            {
                throw new NotImplementedException();
            }
            else
            {
                // Nothing, a fill with no fill/gradient is a valid state per XML
            }

            xml.WriteEndElement();
        }

        xml.WriteEndElement();
    }

    private void WriteBorders(XmlTreeWriter xml, SequentialMap<int, XLBorderFormatValue> idMap)
    {
        xml.WriteStartElement("borders", _ns);
        xml.WriteAttribute("count", idMap.Count);
        foreach (var (_, border) in idMap.GetActual())
        {
            xml.WriteStartElement("border", _ns);
            xml.WriteAttributeDefault("diagonalUp", border.DiagonalUp, false);
            xml.WriteAttributeDefault("diagonalDown", border.DiagonalDown, false);
            xml.WriteAttributeDefault("outline", border.Outline, true);

            // ISO should be "start"+"end", but Excel uses "left"+"right"
            WriteBorderPr("left", border.Left);
            WriteBorderPr("right", border.Right);
            WriteBorderPr("top", border.Top);
            WriteBorderPr("bottom", border.Bottom);
            WriteBorderPr("diagonal", border.Diagonal);
            WriteBorderPr("vertical", border.Vertical);
            WriteBorderPr("horizontal", border.Horizontal);

            xml.WriteEndElement();
        }

        xml.WriteEndElement();
        return;

        void WriteBorderPr(string name, XLBorderLine? borderLine)
        {
            if (!borderLine.HasValue)
                return;

            xml.WriteStartElement(name, _ns);
            xml.WriteAttributeDefault("style", borderLine.Value.Style, XLBorderStyleValues.None);
            xml.WriteColor("color", _ns, borderLine.Value.Color);
            xml.WriteEndElement();
        }
    }

    private void WriteCellXfs(
        XmlTreeWriter xml,
        SequentialMap<int, XLCellFormatValue> idMap,
        SequentialMap<int, string> numFmtIdMap,
        SequentialMap<int, XLFontFormatValue> fontIdMap,
        SequentialMap<int, XLFillFormatValue> fillIdMap,
        SequentialMap<int, XLBorderFormatValue> borderIdMap)
    {
        xml.WriteStartElement("cellXfs", _ns);
        xml.WriteAttribute("count", idMap.Count);
        foreach (var (_, cellXf) in idMap.GetActual())
        {
            xml.WriteStartElement("xf", _ns);

            if (cellXf.NumberFormat is not null)
            {
                if (!XLPredefinedFormat.NumberFormatIds.TryGetValue(cellXf.NumberFormat, out var numFmtId))
                    numFmtId = numFmtIdMap.GetSavedId(cellXf.NumberFormat);

                xml.WriteAttributeOptional("numFmtId", numFmtId);
            }

            if (cellXf.Font is not null)
                xml.WriteAttributeOptional("fontId", fontIdMap.GetSavedId(cellXf.Font));

            if (cellXf.Fill is not null)
                xml.WriteAttributeOptional("fillId", fillIdMap.GetSavedId(cellXf.Fill));

            if (cellXf.Border is not null)
                xml.WriteAttributeOptional("borderId", borderIdMap.GetSavedId(cellXf.Border));

            // TODO: Write cell style
            // xml.WriteAttributeDefault("xfId", cellXf.CellStyleId, 0);
            xml.WriteAttributeDefault("quotePrefix", cellXf.IncludeQuotePrefix, false);
            xml.WriteAttributeDefault("pivotButton", cellXf.PivotButton, false);
            xml.WriteAttributeDefault("applyNumberFormat", cellXf.CustomFormat.HasFlag(CellFormatComponents.NumberFormat), false);
            xml.WriteAttributeDefault("applyFont", cellXf.CustomFormat.HasFlag(CellFormatComponents.Font), false);
            xml.WriteAttributeDefault("applyFill", cellXf.CustomFormat.HasFlag(CellFormatComponents.Fill), false);
            xml.WriteAttributeDefault("applyBorder", cellXf.CustomFormat.HasFlag(CellFormatComponents.Border), false);
            xml.WriteAttributeDefault("applyAlignment", cellXf.CustomFormat.HasFlag(CellFormatComponents.Alignment), false);
            xml.WriteAttributeDefault("applyProtection", cellXf.CustomFormat.HasFlag(CellFormatComponents.Protection), false);

            if (cellXf.Alignment is { } alignment)
            {
                xml.WriteStartElement("alignment", _ns);
                xml.WriteAttributeDefault("horizontal", alignment.Horizontal, XLAlignmentHorizontalValues.General);
                xml.WriteAttributeDefault("vertical", alignment.Vertical, XLAlignmentVerticalValues.Bottom);
                xml.WriteAttributeDefault("textRotation", alignment.TextRotation.GetIso(), 0);
                xml.WriteAttributeDefault("wrapText", alignment.WrapText, false);
                xml.WriteAttributeDefault("indent", alignment.Indent, 0);
                xml.WriteAttributeDefault("relativeIndent", alignment.RelativeIndent, 0);
                xml.WriteAttributeDefault("justifyLastLine", alignment.JustifyLastLine, false);
                xml.WriteAttributeDefault("shrinkToFit", alignment.ShrinkToFit, false);
                xml.WriteAttributeDefault("readingOrder", alignment.ReadingOrder, XLAlignmentReadingOrderValues.ContextDependent);
                xml.WriteEndElement();
            }

            if (cellXf.Protection is { } protection)
            {
                xml.WriteStartElement("protection", _ns);
                xml.WriteAttributeDefault("locked", protection.Locked, true);
                xml.WriteAttributeDefault("hidden", protection.Hidden, false);
                xml.WriteEndElement();
            }

            // TODO: extLst
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
    }

    private class SequentialMap<TKey, T>
        where TKey : struct
    {
        /// <summary>
        /// The index is the one that is used to save the <typeparamref name="T"/> while value is an index to the <c>_fullMap</c>.
        /// </summary>
        private readonly List<TKey> _savedIdToActualId = new();

        private readonly int _offset;

        private readonly IReadOnlyBiDictionary<TKey, T> _fullMap;

        public SequentialMap(IReadOnlyBiDictionary<TKey, T> fullMap, int offset = 0)
        {
            _fullMap = fullMap;
            _offset = offset;
        }

        /// <summary>
        /// How many entries to save are in the map.
        /// </summary>
        public int Count => _savedIdToActualId.Count;

        internal static SequentialMap<TKey, T> Create(HashSet<T> usedValues, IReadOnlyBiDictionary<TKey, T> allValuesMap, int offset = 0)
        {
            var map = new SequentialMap<TKey, T>(allValuesMap, offset);
            foreach (var (actualId, value) in allValuesMap)
            {
                if (!usedValues.Contains(value))
                    continue;

                map.Add(actualId);
            }

            map.Sort();
            return map;
        }

        public void Add(TKey actualId)
        {
            _savedIdToActualId.Add(actualId);
        }

        public void Sort()
        {
            _savedIdToActualId.Sort();
        }

        public IEnumerable<(int SaveId, T Actual)> GetActual()
        {
            return _savedIdToActualId.Select((value, index) => (index + _offset, _fullMap[value]));
        }

        public int GetSavedId(T item)
        {
            var actualId = _fullMap[item];

            // TODO: make another identity map, plus this returns -1 on error
            return _savedIdToActualId.IndexOf(actualId);
        }
    }
}
