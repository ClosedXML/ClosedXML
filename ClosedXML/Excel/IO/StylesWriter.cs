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
using PivotRegion = ClosedXML.Excel.Formatting.XLPivotStyleRegionValues;
using TableRegion = ClosedXML.Excel.Formatting.XLTableStyleRegionValues;

namespace ClosedXML.Excel.IO;

internal class StylesWriter
{
    private const int FirstUserDefinedFormatIndex = XLWorkbookStyles.FirstUserDefinedNumberFormatIndex;

    private static readonly List<(string Type, TableRegion? TableRegion, PivotRegion? PivotRegion)> TableRegionsMap = new()
    {
        ("wholeTable", TableRegion.WholeTable, PivotRegion.WholeTable),
        ("headerRow", TableRegion.HeaderRow, PivotRegion.HeaderRow),
        ("totalRow", TableRegion.TotalRow, PivotRegion.GrandTotalRow),
        ("firstColumn", TableRegion.FirstColumn, PivotRegion.FirstColumn),
        ("lastColumn", TableRegion.LastColumn, PivotRegion.GrandTotalColumn),
        ("firstRowStripe", TableRegion.FirstRowStripe, PivotRegion.FirstRowStripe),
        ("secondRowStripe", TableRegion.SecondRowStripe, PivotRegion.SecondRowStripe),
        ("firstColumnStripe", TableRegion.FirstColumnStripe, PivotRegion.FirstColumnStripe),
        ("secondColumnStripe", TableRegion.SecondColumnStripe, PivotRegion.SecondColumnStripe),
        ("firstHeaderCell", TableRegion.FirstHeaderCell, PivotRegion.FirstHeaderCell),
        ("lastHeaderCell", TableRegion.LastHeaderCell, null),
        ("firstTotalCell", TableRegion.FirstTotalCell, null),
        ("lastTotalCell", TableRegion.LastTotalCell, null),
        ("firstSubtotalColumn", null, PivotRegion.SubtotalColumn1),
        ("secondSubtotalColumn", null, PivotRegion.SubtotalColumn2),
        ("thirdSubtotalColumn", null, PivotRegion.SubtotalColumn3),
        ("firstSubtotalRow", null, PivotRegion.SubtotalRow1),
        ("secondSubtotalRow", null, PivotRegion.SubtotalRow2),
        ("thirdSubtotalRow", null, PivotRegion.SubtotalRow3),
        ("blankRow", null, PivotRegion.BlankRow),
        ("firstColumnSubheading", null, PivotRegion.ColumnSubheading1),
        ("secondColumnSubheading", null, PivotRegion.ColumnSubheading2),
        ("thirdColumnSubheading", null, PivotRegion.ColumnSubheading3),
        ("firstRowSubheading", null, PivotRegion.RowSubheading1),
        ("secondRowSubheading", null, PivotRegion.RowSubheading2),
        ("thirdRowSubheading", null, PivotRegion.RowSubheading3),
        ("pageFieldLabels", null, PivotRegion.PageFieldLabels),
        ("pageFieldValues", null, PivotRegion.PageFieldValues),
    };

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

        foreach (var cellStyle in styles.CellStyles.Values)
        {
            if (cellStyle.NumberFormat is { } numberFormat)
                usedNumberFormats.Add(numberFormat);

            if (cellStyle.Font is { } font)
                usedFonts.Add(font);

            if (cellStyle.Fill is { } fill)
                usedFills.Add(fill);

            if (cellStyle.Border is { } border)
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
        // The map has predefined formats from index 0 and the user defined ones from 164 onward.
        // There is a gap between predefined formats.
        var predefinedNumberFormats = XLPredefinedFormat.FormatCodes.OrderBy(x => x.Key).Select(x => x.Value).ToArray();
        var numberFormatMap = SequentialMap<int, string>.Create(usedNumberFormats, styles.NumberFormats, FirstUserDefinedFormatIndex, predefinedNumberFormats);

        if (numberFormatMap.Count > predefinedNumberFormats.Length)
            WriteNumberFormats(xml, numberFormatMap);

        // Fonts. Register default format font as font zero. The font zero is used for font name and size.
        var fontFormatsMap = SequentialMap<int, XLFontFormatValue>.Create(usedFonts, styles.Fonts, 0, styles.DefaultFormat.Font);
        if (fontFormatsMap.Count > 0)
            WriteFonts(xml, fontFormatsMap);

        // Fill 0 must be None and fill 1 must be Gray125, that is just an immutable fact of the universe.
        // Excel will ignore fills at 0/1 and will use None/Gray125. Write both fills whether they are used
        // or not.
        AddFillAsUsed(XLFillFormatValue.None);
        AddFillAsUsed(XLFillFormatValue.Gray125);
        var fillsFormatsMap = SequentialMap<int, XLFillFormatValue>.Create(usedFills, styles.Fills, 0, XLFillFormatValue.None, XLFillFormatValue.Gray125);
        WriteFills(xml, fillsFormatsMap);

        var borderFormatsMap = SequentialMap<int, XLBorderFormatValue>.Create(usedBorders, styles.Borders);
        if (borderFormatsMap.Count > 0)
            WriteBorders(xml, borderFormatsMap);

        // TODO Styles: Ensure normal style is written, though that should be done during initialization/loading
        var cellStylesMap = new SequentialMap<StyleId, XLCellStyleValue>(styles.CellStyles);

        // All cell styles, regardless if they are used or not should be written to the file
        foreach (var styleId in styles.CellStyles.Keys)
            cellStylesMap.Add(styleId);

        cellStylesMap.Sort();
        if (cellStylesMap.Count > 0)
            WriteCellStyleXfs(xml, cellStylesMap, numberFormatMap, fontFormatsMap, fillsFormatsMap, borderFormatsMap);

        var cellXfsMap = SequentialMap<int, XLCellFormatValue>.Create(usedCellFormats, styles.CellFormats, 0, styles.DefaultFormat);
        if (cellXfsMap.Count > 0)
            WriteCellXfs(xml, cellXfsMap, numberFormatMap, fontFormatsMap, fillsFormatsMap, borderFormatsMap, cellStylesMap);

        if (cellStylesMap.Count > 0)
            WriteCellStyles(xml, cellStylesMap);

        // TODO: Create dxfMap from used dxfs in tables, pivot tables and so on
        var dxfMap = new SequentialMap<int, XLDxfValue>(styles.DifferentialFormats);
        foreach (var dxfId in styles.DifferentialFormats.Keys)
            dxfMap.Add(dxfId);

        dxfMap.Sort();

        WriteDxfs(xml, dxfMap, numberFormatMap.Count);

        var hasTableStyles = styles.TableStyles.Count > 0 ||
                             styles.PivotStyles.Count > 0 ||
                             styles.DefaultTableStyle is not null ||
                             styles.DefaultPivotStyle is not null;
        if (hasTableStyles)
            WriteTableStyles(xml, dxfMap, styles);

        WriteColors(xml, styles);

        xml.WriteEndElement();

        // Fill the map used in other parts to determine format id
        foreach (var (xfId, format) in cellXfsMap.GetActual())
        {
            if (!context.FormatMap.ContainsKey(format))
                context.FormatMap.Add(format, (uint)xfId);
        }

        return;

        void AddFillAsUsed(XLFillFormatValue format)
        {
            if (!styles.Fills.ContainsValue(format))
                styles.AddFillFormat(format);

            usedFills.Add(format);
        }
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

            WriteNumFmt(xml, "numFmt", savedId, format);
        }

        xml.WriteEndElement(); // numFmts
    }

    private void WriteNumFmt(XmlTreeWriter xml, string elementName, int numFmtId, string format)
    {
        xml.WriteStartElement(elementName, _ns);
        xml.WriteAttribute("numFmtId", numFmtId);
        xml.WriteAttribute("formatCode", format);
        xml.WriteEndElement();
    }

    private void WriteFonts(XmlTreeWriter xml, SequentialMap<int, XLFontFormatValue> idMap)
    {
        xml.WriteStartElement("fonts", _ns);
        xml.WriteAttribute("count", idMap.Count);

        foreach (var (_, font) in idMap.GetActual())
            WriteFont(xml, "font", font);

        xml.WriteEndElement();
    }

    private void WriteFont(XmlTreeWriter xml, string elementName, XLFontFormatValue font)
    {
        // Font table and dxf use same XML type, use adapter.
        var dxfAdapter = new XLDifferentialFontValue
        {
            Name = font.Name,
            Size = font.Size,
            Charset = font.Charset,
            Family = font.Family,
            Bold = font.Bold,
            Italic = font.Italic,
            Strikethrough = font.Strikethrough,
            Outline = font.Outline,
            Shadow = font.Shadow,
            Condense = font.Condense,
            Extend = font.Extend,
            Color = font.Color,
            Underline = font.Underline,
            VerticalAlignment = font.VerticalAlignment,
            Scheme = font.Scheme,
        };
        WriteFont(xml, elementName, dxfAdapter);
    }

    private void WriteFont(XmlTreeWriter xml, string elementName, XLDifferentialFontValue font)
    {
        // MS-OI29500 dictates font elements order.
        xml.WriteStartElement(elementName, _ns);

        WriteFlag("b", font.Bold);
        WriteFlag("i", font.Italic);
        WriteFlag("strike", font.Strikethrough);
        WriteFlag("condense", font.Condense);
        WriteFlag("extend", font.Extend);
        WriteFlag("outline", font.Outline);
        WriteFlag("shadow", font.Shadow);

        if (font.Underline is { } underline && underline != XLFontUnderlineValues.None)
        {
            xml.WriteStartElement("u", _ns);
            xml.WriteAttributeDefault("val", underline, XLFontUnderlineValues.Single);
            xml.WriteEndElement();
        }

        if (font.VerticalAlignment is { } verticalAlignment && verticalAlignment != XLFontVerticalTextAlignmentValues.Baseline)
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

        if (font.Family is { } family && family != XLFontFamilyNumberingValues.NotApplicable)
        {
            xml.WriteStartElement("family", _ns);
            xml.WriteAttribute("val", (int)family);
            xml.WriteEndElement();
        }

        if (font.Charset is { } charset && charset != XLFontCharSet.Ansi)
        {
            // Charset is stored as an CT_IntProperty
            xml.WriteStartElement("charset", _ns);
            xml.WriteAttribute("val", (int)charset);
            xml.WriteEndElement();
        }

        if (font.Scheme is { } scheme && scheme != XLFontScheme.None)
        {
            xml.WriteStartElement("scheme", _ns);
            xml.WriteAttribute("val", scheme);
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
        return;

        void WriteFlag(string flagName, bool? flag)
        {
            if (flag == true)
                xml.WriteBooleanProperty(flagName, true, _ns);
        }
    }

    private void WriteFills(XmlTreeWriter xml, SequentialMap<int, XLFillFormatValue> idMap)
    {
        xml.WriteStartElement("fills", _ns);
        xml.WriteAttribute("count", idMap.Count);

        foreach (var (_, fill) in idMap.GetActual())
            WriteFill(xml, "fill", fill, false);

        xml.WriteEndElement();
    }

    private void WriteFill(XmlTreeWriter xml, string elementName, XLFillFormatValue fill, bool isDxf)
    {
        xml.WriteStartElement(elementName, _ns);

        // A fill element with no pattern/gradient is a valid state per XML
        if (fill.Pattern is { } patternFill)
        {
            xml.WriteStartElement("patternFill", _ns);
            xml.WriteAttribute("patternType", patternFill.PatternType);

            var patternColor = patternFill.PatternColor;
            var bgColor = patternFill.BackgroundColor;

            // Fix solid pattern discrepancy. The GUI shows solid fill color in the background
            // color picker, so it would be expected that it is stored in bgColor.
            // * internal structures store solid fill color in the 'IXLFill.BackgroundColor'
            // * For cell format, the 'patternFill' element stores it in the *fgColor*, not bgColor
            // * For dxf, the 'patternFill' stores it in the *bgColor*
            if (!isDxf && patternFill.PatternType == XLFillPatternValues.Solid)
            {
                (patternColor, bgColor) = (bgColor, patternColor);
            }

            if (patternColor.HasValue)
                xml.WriteColor("fgColor", _ns, patternColor);

            if (bgColor.HasValue)
                xml.WriteColor("bgColor", _ns, bgColor);

            xml.WriteEndElement();
        }
        else if (fill.LinearGradient is { } linearGradient)
        {
            // Linear is the default type, so no need to write it
            xml.WriteStartElement("gradientFill", _ns);
            xml.WriteAttributeDefault("degree", linearGradient.Degrees, 0);

            WriteStops(linearGradient.Stops);

            xml.WriteEndElement();
        }
        else if (fill.PathGradient is { } pathGradient)
        {
            xml.WriteStartElement("gradientFill", _ns);
            xml.WriteAttribute("type", "path");
            xml.WriteAttributeDefault("left", pathGradient.InnerLeft.Value, 0);
            xml.WriteAttributeDefault("right", pathGradient.InnerRight.Value, 0);
            xml.WriteAttributeDefault("top", pathGradient.InnerTop.Value, 0);
            xml.WriteAttributeDefault("bottom", pathGradient.InnerBottom.Value, 0);

            WriteStops(pathGradient.Stops);

            xml.WriteEndElement();
        }

        xml.WriteEndElement();
        return;

        void WriteStops(IReadOnlyDictionary<FractionOfOne, XLColor> stops)
        {
            // Excel doesn't care about stop order by positions, but sort anyway
            foreach (var (position, color) in stops.OrderBy(x => x.Key.Value))
            {
                xml.WriteStartElement("stop", _ns);
                xml.WriteAttribute("position", position.Value);
                xml.WriteColor("color", _ns, color);
                xml.WriteEndElement();
            }
        }
    }

    private void WriteBorders(XmlTreeWriter xml, SequentialMap<int, XLBorderFormatValue> idMap)
    {
        xml.WriteStartElement("borders", _ns);
        xml.WriteAttribute("count", idMap.Count);
        foreach (var (_, border) in idMap.GetActual())
            WriteBorder(xml, "border", border, false);

        xml.WriteEndElement();
    }

    private void WriteBorder(XmlTreeWriter xml, string elementName, XLBorderFormatValue border, bool isDxf)
    {
        xml.WriteStartElement(elementName, _ns);
        xml.WriteAttributeDefault("diagonalUp", border.DiagonalUp, false);
        xml.WriteAttributeDefault("diagonalDown", border.DiagonalDown, false);

        // Outline has no meaning for cell styles, it is only for dxf - tables and such.
        if (isDxf)
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
        return;

        void WriteBorderPr(string name, XLBorderLine? borderLine)
        {
            if (!borderLine.HasValue)
                return;

            xml.WriteStartElement(name, _ns);
            xml.WriteAttributeDefault("style", borderLine.Value.Style, XLBorderStyleValues.None);
            if (borderLine.Value.Style != XLBorderStyleValues.None)
                xml.WriteColor("color", _ns, borderLine.Value.Color);

            xml.WriteEndElement();
        }
    }

    private void WriteCellStyleXfs(
        XmlTreeWriter xml,
        SequentialMap<StyleId, XLCellStyleValue> cellStylesMap,
        SequentialMap<int, string> numFmtIdMap,
        SequentialMap<int, XLFontFormatValue> fontIdMap,
        SequentialMap<int, XLFillFormatValue> fillIdMap,
        SequentialMap<int, XLBorderFormatValue> borderIdMap)
    {
        xml.WriteStartElement("cellStyleXfs", _ns);
        xml.WriteAttribute("count", cellStylesMap.Count);

        // Collection must have at least one element
        foreach (var (_, cellStyle) in cellStylesMap.GetActual())
        {
            xml.WriteStartElement("xf", _ns);
            xml.WriteAttributeOptional("numFmtId", numFmtIdMap.GetSavedId(cellStyle.NumberFormat));
            xml.WriteAttributeOptional("fontId", fontIdMap.GetSavedId(cellStyle.Font));
            xml.WriteAttributeOptional("fillId", fillIdMap.GetSavedId(cellStyle.Fill));
            xml.WriteAttributeOptional("borderId", borderIdMap.GetSavedId(cellStyle.Border));

            // cellStyleXf doesn't use quote, pivot button or xfId -> skip those attributes
            xml.WriteAttributeDefault("applyNumberFormat", cellStyle.IncludedComponents.HasFlag(CellFormatComponents.NumberFormat), true);
            xml.WriteAttributeDefault("applyFont", cellStyle.IncludedComponents.HasFlag(CellFormatComponents.Font), true);
            xml.WriteAttributeDefault("applyFill", cellStyle.IncludedComponents.HasFlag(CellFormatComponents.Fill), true);
            xml.WriteAttributeDefault("applyBorder", cellStyle.IncludedComponents.HasFlag(CellFormatComponents.Border), true);
            xml.WriteAttributeDefault("applyAlignment", cellStyle.IncludedComponents.HasFlag(CellFormatComponents.Alignment), true);
            xml.WriteAttributeDefault("applyProtection", cellStyle.IncludedComponents.HasFlag(CellFormatComponents.Protection), true);

            if (cellStyle.Alignment is { } alignment)
                WriteAlignment(xml, "alignment", alignment);

            if (cellStyle.Protection is { } protection)
                WriteProtection(xml, "protection", protection);

            // TODO: extLst
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
    }

    private void WriteCellXfs(
        XmlTreeWriter xml,
        SequentialMap<int, XLCellFormatValue> idMap,
        SequentialMap<int, string> numFmtIdMap,
        SequentialMap<int, XLFontFormatValue> fontIdMap,
        SequentialMap<int, XLFillFormatValue> fillIdMap,
        SequentialMap<int, XLBorderFormatValue> borderIdMap,
        SequentialMap<StyleId, XLCellStyleValue> cellStyleIdMap)
    {
        xml.WriteStartElement("cellXfs", _ns);
        xml.WriteAttribute("count", idMap.Count);
        foreach (var (_, cellXf) in idMap.GetActual())
        {
            xml.WriteStartElement("xf", _ns);

            xml.WriteAttributeOptional("numFmtId", numFmtIdMap.GetSavedId(cellXf.NumberFormat));
            xml.WriteAttributeOptional("fontId", fontIdMap.GetSavedId(cellXf.Font));
            xml.WriteAttributeOptional("fillId", fillIdMap.GetSavedId(cellXf.Fill));
            xml.WriteAttributeOptional("borderId", borderIdMap.GetSavedId(cellXf.Border));

            if (cellXf.CellStyleId is not null)
                xml.WriteAttributeDefault("xfId", cellStyleIdMap.GetSavedId(cellXf.CellStyleId.Value), 0);

            xml.WriteAttributeDefault("quotePrefix", cellXf.IncludeQuotePrefix, false);
            xml.WriteAttributeDefault("pivotButton", cellXf.PivotButton, false);
            xml.WriteAttributeDefault("applyNumberFormat", cellXf.CustomFormat.HasFlag(CellFormatComponents.NumberFormat), false);
            xml.WriteAttributeDefault("applyFont", cellXf.CustomFormat.HasFlag(CellFormatComponents.Font), false);
            xml.WriteAttributeDefault("applyFill", cellXf.CustomFormat.HasFlag(CellFormatComponents.Fill), false);
            xml.WriteAttributeDefault("applyBorder", cellXf.CustomFormat.HasFlag(CellFormatComponents.Border), false);
            xml.WriteAttributeDefault("applyAlignment", cellXf.CustomFormat.HasFlag(CellFormatComponents.Alignment), false);
            xml.WriteAttributeDefault("applyProtection", cellXf.CustomFormat.HasFlag(CellFormatComponents.Protection), false);

            if (cellXf.Alignment is { } alignment)
                WriteAlignment(xml, "alignment", alignment);

            if (cellXf.Protection is { } protection)
                WriteProtection(xml, "protection", protection);

            // TODO: extLst
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
    }

    private void WriteAlignment(XmlTreeWriter xml, string elementName, XLAlignmentFormatValue alignment)
    {
        xml.WriteStartElement(elementName, _ns);
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

    private void WriteProtection(XmlTreeWriter xml, string elementName, XLProtectionFormatValue protection)
    {
        if (protection.Locked && !protection.Hidden)
            return;

        xml.WriteStartElement(elementName, _ns);
        xml.WriteAttributeDefault("locked", protection.Locked, true);
        xml.WriteAttributeDefault("hidden", protection.Hidden, false);
        xml.WriteEndElement();
    }

    private void WriteCellStyles(XmlTreeWriter xml, SequentialMap<StyleId, XLCellStyleValue> cellStylesMap)
    {
        xml.WriteStartElement("cellStyles", _ns);
        xml.WriteAttribute("count", cellStylesMap.Count);

        // Collection must have at least one element
        foreach (var (mappedStyleId, cellStyle) in cellStylesMap.GetActual())
        {
            xml.WriteStartElement("cellStyle", _ns);

            // Name is technically optional and Excel will generate one if missing, but we ensure the name always exist
            xml.WriteAttribute("name", cellStyle.Name);
            xml.WriteAttribute("xfId", mappedStyleId);
            if (cellStyle.BuiltInStyle is { } builtInStyle)
            {
                if (builtInStyle is >= BuiltInStyleValues.RowLevel1 and <= BuiltInStyleValues.RowLevel7)
                {
                    xml.WriteAttributeOptional("builtinId", 1);
                    xml.WriteAttribute("iLevel", BuiltInStyleValues.RowLevel1 - builtInStyle + 1);
                }
                else if (builtInStyle is >= BuiltInStyleValues.ColumnLevel1 and <= BuiltInStyleValues.ColumnLevel7)
                {
                    xml.WriteAttributeOptional("builtinId", 2);
                    xml.WriteAttribute("iLevel", BuiltInStyleValues.ColumnLevel1 - builtInStyle + 1);
                }
                else
                {
                    xml.WriteAttributeOptional("builtinId", (int?)cellStyle.BuiltInStyle);
                }
            }

            // Hidden + flag are optional per schema, but basically it's a bool with default
            xml.WriteAttributeDefault("hidden", cellStyle.Hidden, false);
            xml.WriteAttributeDefault("customBuiltin", cellStyle.BuiltInStyle is not null, true);
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
    }

    private void WriteDxfs(XmlTreeWriter xml, SequentialMap<int, XLDxfValue> differentialFormats, int lastNumFmtId)
    {
        xml.WriteStartElement("dxfs", _ns);
        xml.WriteAttribute("count", differentialFormats.Count);
        foreach (var (_, dxf) in differentialFormats.GetActual())
        {
            xml.WriteStartElement("dxf", _ns);

            if (dxf.Font is { } font)
                WriteFont(xml, "font", font);

            if (dxf.NumberFormat is { } numberFormat)
            {
                // numFmtId doesn't matter in dxf, but keep them unique (Excel-like behavior)
                WriteNumFmt(xml, "numFmt", ++lastNumFmtId, numberFormat);
            }

            if (dxf.Fill is { } fill)
                WriteFill(xml, "fill", fill, true);

            if (dxf.Alignment is { } alignment)
                WriteAlignment(xml, "alignment", alignment);

            if (dxf.Border is { } border)
                WriteBorder(xml, "border", border, true);

            if (dxf.Protection is { } protection)
                WriteProtection(xml, "protection", protection);

            // TODO: extLst
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
    }

    private void WriteTableStyles(XmlTreeWriter xml, SequentialMap<int, XLDxfValue> dxfMap, XLWorkbookStyles styles)
    {
        var allStyleNames = styles.TableStyles.Keys.Concat(styles.PivotStyles.Keys).Distinct(XLHelper.NameComparer).ToList();
        allStyleNames.Sort();

        xml.WriteStartElement("tableStyles", _ns);
        xml.WriteAttribute("count", allStyleNames.Count);
        if (styles.DefaultTableStyle is { } defaultTableStyle)
            xml.WriteAttribute("defaultTableStyle", defaultTableStyle);

        if (styles.DefaultPivotStyle is { } defaultPivotStyle)
            xml.WriteAttribute("defaultPivotStyle", defaultPivotStyle);

        foreach (var styleName in allStyleNames)
            WriteTableStyle(xml, styleName, dxfMap, styles);

        xml.WriteEndElement();
    }

    private void WriteTableStyle(XmlTreeWriter xml, string styleName, SequentialMap<int, XLDxfValue> dxfMap, XLWorkbookStyles styles)
    {
        xml.WriteStartElement("tableStyle", _ns);
        xml.WriteAttribute("name", styleName);

        var hasPivotStyle = styles.PivotStyles.TryGetValue(styleName, out var pivotStyle);
        xml.WriteAttributeDefault("pivot", hasPivotStyle, true);

        var hasTableStyle = styles.TableStyles.TryGetValue(styleName, out var tableStyle);
        xml.WriteAttributeDefault("table", hasTableStyle, true);

        var styledRegions = new List<(string Type, int Size, int DxfId)>(TableRegionsMap.Count);
        foreach (var (type, tableRegion, pivotRegion) in TableRegionsMap)
        {
            var tableRegionStyle = TryGetTableRegion(tableStyle, tableRegion);
            var pivotRegionStyle = TryGetPivotRegion(pivotStyle, pivotRegion);
            if (tableRegionStyle is var (tableSize, tableDxf) &&
                pivotRegionStyle is var (pivotSize, pivotDxf) &&
                (tableSize != pivotSize || tableDxf != pivotDxf))
            {
                // This should never happen. The table/pivot shared style have same band size and
                // dxf on load and we don't provide API to modify table/pivot styles.
                // Sidenote: Excel GUI will refuse to create table style that conflicts with a pivot
                // style and vice versa. It will show an alert 'This style name already exists.'
                throw new InvalidOperationException($"Table and pivot table style '{styleName}' that has different formatting for {tableRegion}/{pivotRegion}.");
            }

            if ((tableRegionStyle ?? pivotRegionStyle) is var (bandSize, dxf))
                styledRegions.Add((type, bandSize, dxfMap.GetSavedId(dxf)));
        }

        xml.WriteAttribute("count", styledRegions.Count);
        foreach (var (type, bandSize, dxfId) in styledRegions)
        {
            xml.WriteStartElement("tableStyleElement", _ns);
            xml.WriteAttribute("type", type);
            xml.WriteAttributeDefault("size", bandSize, 1);
            xml.WriteAttribute("dxfId", dxfId);
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
        return;

        static (int Size, XLDxfValue Dxf)? TryGetTableRegion(XLTableTheme? tableStyle, TableRegion? tableRegion)
        {
            if (tableStyle is null || tableRegion is null)
                return null;

            if (!tableStyle.RegionFormats.TryGetValue(tableRegion.Value, out var dxf))
                return null;

            var bandSize = tableRegion switch
            {
                TableRegion.FirstRowStripe => tableStyle.RowStripe1BandSize,
                TableRegion.SecondRowStripe => tableStyle.RowStripe2BandSize,
                TableRegion.FirstColumnStripe => tableStyle.ColumnStripe1BandSize,
                TableRegion.SecondColumnStripe => tableStyle.ColumnStripe2BandSize,
                _ => 1
            };
            return (bandSize, dxf);
        }

        static (int Size, XLDxfValue Dxf)? TryGetPivotRegion(XLPivotTableStyle? pivotStyle, PivotRegion? region)
        {
            if (pivotStyle is null || region is null)
                return null;

            if (!pivotStyle.RegionFormats.TryGetValue(region.Value, out var dxf))
                return null;

            var bandSize = region switch
            {
                PivotRegion.FirstRowStripe => pivotStyle.RowStripe1BandSize,
                PivotRegion.SecondRowStripe => pivotStyle.RowStripe2BandSize,
                PivotRegion.FirstColumnStripe => pivotStyle.ColumnStripe1BandSize,
                PivotRegion.SecondColumnStripe => pivotStyle.ColumnStripe2BandSize,
                _ => 1
            };
            return (bandSize, dxf);
        }
    }

    private void WriteColors(XmlTreeWriter xml, XLWorkbookStyles styles)
    {
        var hasMruColors = styles.MruColors.Count > 0;
        var indexedColors = styles.IndexedColorsArgb;
        var hasIndexColors = indexedColors is { Count: > 0 };
        if (!hasMruColors && !hasIndexColors)
            return;

        xml.WriteStartElement("colors", _ns);

        if (hasIndexColors)
        {
            xml.WriteStartElement("indexedColors", _ns);
            foreach (var indexedColor in indexedColors!)
            {
                xml.WriteStartElement("rgbColor", _ns);
                xml.WriteAttributeHex("rgb", indexedColor);
                xml.WriteEndElement();
            }

            xml.WriteEndElement();
        }

        if (hasMruColors)
        {
            xml.WriteStartElement("mruColors", _ns);
            foreach (var mruColor in styles.MruColors)
                xml.WriteColor("color", _ns, mruColor);

            xml.WriteEndElement();
        }

        xml.WriteEndElement(); // colors
    }

    private class SequentialMap<TKey, T>
        where TKey : struct
    {
        /// <summary>
        /// The index is the one that is used to save the <typeparamref name="T"/> while value is an index to the <c>_fullMap</c>.
        /// </summary>
        private readonly Dictionary<int, TKey> _savedIdToActualId = new();

        private readonly IReadOnlyBiDictionary<TKey, T> _fullMap;

        /// <summary>
        /// A table that will be saved to file. Contains used and necessary entries along with
        /// the id under which the entry can be retrieved.
        /// </summary>
        private List<(int SaveId, T Actual)>? _saveTable;

        public SequentialMap(IReadOnlyBiDictionary<TKey, T> fullMap)
        {
            _fullMap = fullMap;
        }

        /// <summary>
        /// How many entries to save are in the map.
        /// </summary>
        public int Count => _savedIdToActualId.Count;

        internal static SequentialMap<TKey, T> Create(HashSet<T> usedValues, IReadOnlyBiDictionary<TKey, T> allValuesMap, int usedStart = 0, params T[] firstValues)
        {
            var map = new SequentialMap<TKey, T>(allValuesMap);
            foreach (var firstValue in firstValues)
            {
                var actualId = allValuesMap[firstValue];
                map.Add(actualId);
            }

            // This is here basically for number formats. It ensures that user defined number
            // formats start at 164 and the 0-164 is reserved for predefined formats.
            // Number formats is the only table that can have gaps in the ids.
            var usedSaveId = Math.Max(map.Count, usedStart);
            foreach (var (actualId, value) in allValuesMap)
            {
                if (firstValues.Contains(value))
                    continue;

                if (!usedValues.Contains(value))
                    continue;

                map.Add(actualId, usedSaveId++);
            }

            map.Sort();
            return map;
        }

        public void Add(TKey actualId)
        {
            _savedIdToActualId.Add(_savedIdToActualId.Count, actualId);
        }

        private void Add(TKey actualId, int saveId)
        {
            _savedIdToActualId.Add(saveId, actualId);
        }

        public void Sort()
        {
            _saveTable = _savedIdToActualId
                .Select(x => (x.Key, _fullMap[x.Value]))
                .OrderBy(x => x.Item1)
                .ToList();
        }

        public IEnumerable<(int SaveId, T Actual)> GetActual()
        {
            return _saveTable!;
        }

        public int GetSavedId(T item)
        {
            var actualId = _fullMap[item];
            return GetSavedId(actualId);
        }

        public int GetSavedId(TKey actualId)
        {
            // TODO Styles: Use a better better internal structure
            foreach (var (mapSaveId, mapActualId) in _savedIdToActualId)
            {
                if (mapActualId.Equals(actualId))
                    return mapSaveId;
            }

            throw new InvalidOperationException($"Unable to find saveId for {actualId} of {typeof(T).Name}.");
        }
    }
}
