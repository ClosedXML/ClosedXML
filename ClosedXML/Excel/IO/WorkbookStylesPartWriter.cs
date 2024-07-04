#nullable disable

using ClosedXML.Utils;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using static ClosedXML.Excel.XLWorkbook;

namespace ClosedXML.Excel.IO
{
    internal class WorkbookStylesPartWriter
    {
        internal static void GenerateContent(WorkbookStylesPart workbookStylesPart, XLWorkbook workbook, SaveContext context)
        {
            var defaultStyle = DefaultStyleValue;

            if (!context.SharedFonts.ContainsKey(defaultStyle.Font))
                context.SharedFonts.Add(defaultStyle.Font, new FontInfo { FontId = 0, Font = defaultStyle.Font });

            if (workbookStylesPart.Stylesheet == null)
                workbookStylesPart.Stylesheet = new Stylesheet();

            // Cell styles = Named styles
            if (workbookStylesPart.Stylesheet.CellStyles == null)
                workbookStylesPart.Stylesheet.CellStyles = new CellStyles();

            // To determine the default workbook style, we look for the style with builtInId = 0 (I hope that is the correct approach)
            UInt32 defaultFormatId;
            if (workbookStylesPart.Stylesheet.CellStyles.Elements<CellStyle>().Any(c => c.BuiltinId != null && c.BuiltinId.HasValue && c.BuiltinId.Value == 0))
            {
                // Possible to have duplicate default cell styles - occurs when file gets saved under different cultures.
                // We prefer the style that is named Normal
                var normalCellStyles = workbookStylesPart.Stylesheet.CellStyles.Elements<CellStyle>()
                    .Where(c => c.BuiltinId != null && c.BuiltinId.HasValue && c.BuiltinId.Value == 0)
                    .OrderBy(c => c.Name != null && c.Name.HasValue && c.Name.Value == "Normal");

                defaultFormatId = normalCellStyles.Last().FormatId.Value;
            }
            else if (workbookStylesPart.Stylesheet.CellStyles.Elements<CellStyle>().Any())
                defaultFormatId = workbookStylesPart.Stylesheet.CellStyles.Elements<CellStyle>().Max(c => c.FormatId.Value) + 1;
            else
                defaultFormatId = 0;

            context.SharedStyles.Add(defaultStyle,
                new StyleInfo
                {
                    StyleId = defaultFormatId,
                    Style = defaultStyle,
                    FontId = 0,
                    FillId = 0,
                    BorderId = 0,
                    IncludeQuotePrefix = false,
                    NumberFormatId = 0
                    //AlignmentId = 0
                });

            UInt32 styleCount = 1;
            UInt32 fontCount = 1;
            UInt32 fillCount = 3;
            UInt32 borderCount = 1;
            var xlPivotTablesCustomFormats = new HashSet<string>();
            var xlStyles = new HashSet<XLStyleValue>();

            foreach (var worksheet in workbook.WorksheetsInternal)
            {
                xlStyles.Add(worksheet.StyleValue);
                foreach (var s in worksheet.Internals.ColumnsCollection.Select(c => c.Value.StyleValue))
                {
                    xlStyles.Add(s);
                }
                foreach (var s in worksheet.Internals.RowsCollection.Select(r => r.Value.StyleValue))
                {
                    xlStyles.Add(s);
                }

                foreach (var c in worksheet.Internals.CellsCollection.GetCells())
                {
                    xlStyles.Add(c.StyleValue);
                }

                var xlPivotTableCustomFormats = worksheet.PivotTables
                    .SelectMany<XLPivotTable, XLPivotDataField>(pt => pt.DataFields)
                    .Where(x => x.NumberFormatValue is not null && !string.IsNullOrEmpty(x.NumberFormatValue.Format))
                    .Select(x => x.NumberFormatValue.Format);
                xlPivotTablesCustomFormats.UnionWith(xlPivotTableCustomFormats);
            }

            var alignments = xlStyles.Select(s => s.Alignment).Distinct().ToList();
            var borders = xlStyles.Select(s => s.Border).Distinct().ToList();
            var fonts = xlStyles.Select(s => s.Font).Distinct().ToList();
            var fills = xlStyles.Select(s => s.Fill).Distinct().ToList();
            var numberFormats = xlStyles.Select(s => s.NumberFormat).Distinct().ToList();
            var protections = xlStyles.Select(s => s.Protection).Distinct().ToList();

            for (int i = 0; i < fonts.Count; i++)
            {
                if (!context.SharedFonts.ContainsKey(fonts[i]))
                {
                    context.SharedFonts.Add(fonts[i], new FontInfo { FontId = (uint)fontCount++, Font = fonts[i] });
                }
            }

            var sharedFills = fills.ToDictionary(
                f => f, f => new FillInfo { FillId = fillCount++, Fill = f });

            var sharedBorders = borders.ToDictionary(
                b => b, b => new BorderInfo { BorderId = borderCount++, Border = b });

            var customNumberFormats = numberFormats
                .Where(nf => nf.NumberFormatId == -1)
                .ToHashSet();

            foreach (var pivotNumberFormat in xlPivotTablesCustomFormats)
            {
                var numberFormatKey = XLNumberFormatKey.ForFormat(pivotNumberFormat);
                var numberFormat = XLNumberFormatValue.FromKey(ref numberFormatKey);

                customNumberFormats.Add(numberFormat);
            }

            var allSharedNumberFormats = ResolveNumberFormats(workbookStylesPart, customNumberFormats, defaultFormatId);
            foreach (var nf in allSharedNumberFormats)
            {
                context.SharedNumberFormats.Add(nf.Key, nf.Value);
            }

            ResolveFonts(workbookStylesPart, context);
            var allSharedFills = ResolveFills(workbookStylesPart, sharedFills);
            var allSharedBorders = ResolveBorders(workbookStylesPart, sharedBorders);

            foreach (var xlStyle in xlStyles)
            {
                var numberFormatId = xlStyle.NumberFormat.NumberFormatId >= 0
                    ? xlStyle.NumberFormat.NumberFormatId
                    : allSharedNumberFormats[xlStyle.NumberFormat].NumberFormatId;

                if (!context.SharedStyles.ContainsKey(xlStyle))
                    context.SharedStyles.Add(xlStyle,
                        new StyleInfo
                        {
                            StyleId = styleCount++,
                            Style = xlStyle,
                            FontId = context.SharedFonts[xlStyle.Font].FontId,
                            FillId = allSharedFills[xlStyle.Fill].FillId,
                            BorderId = allSharedBorders[xlStyle.Border].BorderId,
                            NumberFormatId = numberFormatId,
                            IncludeQuotePrefix = xlStyle.IncludeQuotePrefix
                        });
            }

            ResolveCellStyleFormats(workbookStylesPart, context);
            ResolveRest(workbookStylesPart, context);

            if (!workbookStylesPart.Stylesheet.CellStyles.Elements<CellStyle>().Any(c => c.BuiltinId != null && c.BuiltinId.HasValue && c.BuiltinId.Value == 0U))
                workbookStylesPart.Stylesheet.CellStyles.AppendChild(new CellStyle { Name = "Normal", FormatId = defaultFormatId, BuiltinId = 0U });

            workbookStylesPart.Stylesheet.CellStyles.Count = (UInt32)workbookStylesPart.Stylesheet.CellStyles.Count();

            var newSharedStyles = new Dictionary<XLStyleValue, StyleInfo>();
            foreach (var ss in context.SharedStyles)
            {
                var styleId = -1;
                foreach (CellFormat f in workbookStylesPart.Stylesheet.CellFormats)
                {
                    styleId++;
                    if (CellFormatsAreEqual(f, ss.Value, compareAlignment: true))
                        break;
                }
                if (styleId == -1)
                    styleId = 0;
                var si = ss.Value;
                si.StyleId = (UInt32)styleId;
                newSharedStyles.Add(ss.Key, si);
            }
            context.SharedStyles.Clear();
            newSharedStyles.ForEach(kp => context.SharedStyles.Add(kp.Key, kp.Value));

            AddDifferentialFormats(workbookStylesPart, workbook, context);
        }

        /// <summary>
        /// Populates the differential formats that are currently in the file to the SaveContext
        /// </summary>
        private static void AddDifferentialFormats(WorkbookStylesPart workbookStylesPart, XLWorkbook workbook, SaveContext context)
        {
            if (workbookStylesPart.Stylesheet.DifferentialFormats == null)
                workbookStylesPart.Stylesheet.DifferentialFormats = new DifferentialFormats();

            var differentialFormats = workbookStylesPart.Stylesheet.DifferentialFormats;
            differentialFormats.RemoveAllChildren();
            FillDifferentialFormatsCollection(differentialFormats, context.DifferentialFormats);

            foreach (var ws in workbook.WorksheetsInternal)
            {
                foreach (var cf in ws.ConditionalFormats)
                {
                    var styleValue = (cf.Style as XLStyle).Value;
                    if (!styleValue.Equals(DefaultStyleValue) && !context.DifferentialFormats.ContainsKey(styleValue))
                        AddConditionalDifferentialFormat(workbookStylesPart.Stylesheet.DifferentialFormats, cf, context);
                }

                foreach (var tf in ws.Tables.SelectMany<XLTable, IXLTableField>(t => t.Fields))
                {
                    if (tf.IsConsistentStyle())
                    {
                        var style = (tf.Column.Cells()
                            .Skip(tf.Table.ShowHeaderRow ? 1 : 0)
                            .First()
                            .Style as XLStyle).Value;

                        if (!style.Equals(DefaultStyleValue) && !context.DifferentialFormats.ContainsKey(style))
                            AddStyleAsDifferentialFormat(workbookStylesPart.Stylesheet.DifferentialFormats, style, context);
                    }
                }

                foreach (var pt in ws.PivotTables)
                {
                    // Add Dxf from old pivot table structures.
                    foreach (var styleFormat in pt.AllStyleFormats)
                    {
                        var xlStyle = (XLStyle)styleFormat.Style;
                        if (!xlStyle.Value.Equals(DefaultStyleValue) && !context.DifferentialFormats.ContainsKey(xlStyle.Value))
                            AddStyleAsDifferentialFormat(workbookStylesPart.Stylesheet.DifferentialFormats, xlStyle.Value, context);
                    }

                    // Add Dxf from new pivot table structures.
                    foreach (var xlPivotFormat in pt.Formats)
                    {
                        var xlStyle = xlPivotFormat.DxfStyle;
                        if (xlStyle.Value != XLStyleValue.Default &&
                            !context.DifferentialFormats.ContainsKey(xlStyle.Value))
                        {
                            AddStyleAsDifferentialFormat(workbookStylesPart.Stylesheet.DifferentialFormats, xlStyle.Value, context);
                        }
                    }

                    foreach (var xlConditionalStyle in pt.ConditionalFormats)
                    {
                        var xlStyle = ((XLStyle)xlConditionalStyle.Format.Style);
                        if (xlStyle.Value != XLStyleValue.Default &&
                            !context.DifferentialFormats.ContainsKey(xlStyle.Value))
                        {
                            AddStyleAsDifferentialFormat(workbookStylesPart.Stylesheet.DifferentialFormats, xlStyle.Value, context);
                        }
                    }
                }
            }

            differentialFormats.Count = (UInt32)differentialFormats.Count();
            if (differentialFormats.Count == 0)
                workbookStylesPart.Stylesheet.DifferentialFormats = null;
        }

        private static void FillDifferentialFormatsCollection(DifferentialFormats differentialFormats,
            Dictionary<XLStyleValue, int> dictionary)
        {
            dictionary.Clear();
            var id = 0;

            foreach (var df in differentialFormats.Elements<DifferentialFormat>())
            {
                var emptyContainer = new XLStylizedEmpty(DefaultStyle);

                var style = new XLStyle(emptyContainer, DefaultStyle);
                OpenXmlHelper.LoadFont(df.Font, emptyContainer.Style.Font);
                OpenXmlHelper.LoadBorder(df.Border, emptyContainer.Style.Border);
                OpenXmlHelper.LoadNumberFormat(df.NumberingFormat, emptyContainer.Style.NumberFormat);
                OpenXmlHelper.LoadFill(df.Fill, emptyContainer.Style.Fill, differentialFillFormat: true);

                if (!dictionary.ContainsKey(emptyContainer.StyleValue))
                    dictionary.Add(emptyContainer.StyleValue, id++);
            }
        }

        private static void AddConditionalDifferentialFormat(DifferentialFormats differentialFormats, IXLConditionalFormat cf,
            SaveContext context)
        {
            var differentialFormat = new DifferentialFormat();
            var styleValue = (cf.Style as XLStyle).Value;

            var diffFont = GetNewFont(new FontInfo { Font = styleValue.Font }, false);
            if (diffFont?.HasChildren ?? false)
                differentialFormat.Append(diffFont);

            if (!String.IsNullOrWhiteSpace(cf.Style.NumberFormat.Format))
            {
                var numberFormat = new NumberingFormat
                {
                    NumberFormatId = (UInt32)(XLConstants.NumberOfBuiltInStyles + differentialFormats.Count()),
                    FormatCode = cf.Style.NumberFormat.Format
                };
                differentialFormat.Append(numberFormat);
            }

            var diffFill = GetNewFill(new FillInfo { Fill = styleValue.Fill }, differentialFillFormat: true, ignoreMod: false);
            if (diffFill?.HasChildren ?? false)
                differentialFormat.Append(diffFill);

            var diffBorder = GetNewBorder(new BorderInfo { Border = styleValue.Border }, false);
            if (diffBorder?.HasChildren ?? false)
                differentialFormat.Append(diffBorder);

            differentialFormats.Append(differentialFormat);

            context.DifferentialFormats.Add(styleValue, differentialFormats.Count() - 1);
        }

        private static void AddStyleAsDifferentialFormat(DifferentialFormats differentialFormats, XLStyleValue style,
            SaveContext context)
        {
            var differentialFormat = new DifferentialFormat();

            var diffFont = GetNewFont(new FontInfo { Font = style.Font }, false);
            if (diffFont?.HasChildren ?? false)
                differentialFormat.Append(diffFont);

            if (!String.IsNullOrWhiteSpace(style.NumberFormat.Format) || style.NumberFormat.NumberFormatId != 0)
            {
                var numberFormat = new NumberingFormat();

                if (style.NumberFormat.NumberFormatId == -1)
                {
                    numberFormat.FormatCode = style.NumberFormat.Format;
                    numberFormat.NumberFormatId = (UInt32)(XLConstants.NumberOfBuiltInStyles +
                        differentialFormats
                            .Descendants<DifferentialFormat>()
                            .Count(df => df.NumberingFormat != null && df.NumberingFormat.NumberFormatId != null && df.NumberingFormat.NumberFormatId.Value >= XLConstants.NumberOfBuiltInStyles));
                }
                else
                {
                    numberFormat.NumberFormatId = (UInt32)(style.NumberFormat.NumberFormatId);
                    if (!string.IsNullOrEmpty(style.NumberFormat.Format))
                        numberFormat.FormatCode = style.NumberFormat.Format;
                    else if (XLPredefinedFormat.FormatCodes.TryGetValue(style.NumberFormat.NumberFormatId, out string formatCode))
                        numberFormat.FormatCode = formatCode;
                }

                differentialFormat.Append(numberFormat);
            }

            var diffFill = GetNewFill(new FillInfo { Fill = style.Fill }, differentialFillFormat: true, ignoreMod: false);
            if (diffFill?.HasChildren ?? false)
                differentialFormat.Append(diffFill);

            var diffBorder = GetNewBorder(new BorderInfo { Border = style.Border }, false);
            if (diffBorder?.HasChildren ?? false)
                differentialFormat.Append(diffBorder);

            differentialFormats.Append(differentialFormat);

            context.DifferentialFormats.Add(style, differentialFormats.Count() - 1);
        }

        private static void ResolveRest(WorkbookStylesPart workbookStylesPart, SaveContext context)
        {
            if (workbookStylesPart.Stylesheet.CellFormats == null)
                workbookStylesPart.Stylesheet.CellFormats = new CellFormats();

            foreach (var styleInfo in context.SharedStyles.Values)
            {
                var info = styleInfo;
                var foundOne =
                    workbookStylesPart.Stylesheet.CellFormats.Cast<CellFormat>().Any(f => CellFormatsAreEqual(f, info, compareAlignment: true));

                if (foundOne) continue;

                var cellFormat = GetCellFormat(styleInfo);
                cellFormat.FormatId = 0;
                var alignment = new Alignment
                {
                    Horizontal = styleInfo.Style.Alignment.Horizontal.ToOpenXml(),
                    Vertical = styleInfo.Style.Alignment.Vertical.ToOpenXml(),
                    Indent = (UInt32)styleInfo.Style.Alignment.Indent,
                    ReadingOrder = (UInt32)styleInfo.Style.Alignment.ReadingOrder,
                    WrapText = styleInfo.Style.Alignment.WrapText,
                    TextRotation = (UInt32)GetOpenXmlTextRotation(styleInfo.Style.Alignment),
                    ShrinkToFit = styleInfo.Style.Alignment.ShrinkToFit,
                    RelativeIndent = styleInfo.Style.Alignment.RelativeIndent,
                    JustifyLastLine = styleInfo.Style.Alignment.JustifyLastLine
                };
                cellFormat.AppendChild(alignment);

                if (cellFormat.ApplyProtection.Value)
                    cellFormat.AppendChild(GetProtection(styleInfo));

                workbookStylesPart.Stylesheet.CellFormats.AppendChild(cellFormat);
            }
            workbookStylesPart.Stylesheet.CellFormats.Count = (UInt32)workbookStylesPart.Stylesheet.CellFormats.Count();

            static int GetOpenXmlTextRotation(XLAlignmentValue alignment)
            {
                var textRotation = alignment.TextRotation;
                return textRotation >= 0
                    ? textRotation
                    : 90 - textRotation;
            }
        }

        private static void ResolveCellStyleFormats(WorkbookStylesPart workbookStylesPart,
            SaveContext context)
        {
            if (workbookStylesPart.Stylesheet.CellStyleFormats == null)
                workbookStylesPart.Stylesheet.CellStyleFormats = new CellStyleFormats();

            foreach (var styleInfo in context.SharedStyles.Values)
            {
                var info = styleInfo;
                var foundOne =
                    workbookStylesPart.Stylesheet.CellStyleFormats.Cast<CellFormat>().Any(
                        f => CellFormatsAreEqual(f, info, compareAlignment: false));

                if (foundOne) continue;

                var cellStyleFormat = GetCellFormat(styleInfo);

                if (cellStyleFormat.ApplyProtection.Value)
                    cellStyleFormat.AppendChild(GetProtection(styleInfo));

                workbookStylesPart.Stylesheet.CellStyleFormats.AppendChild(cellStyleFormat);
            }
            workbookStylesPart.Stylesheet.CellStyleFormats.Count =
                (UInt32)workbookStylesPart.Stylesheet.CellStyleFormats.Count();
        }

        private static bool ApplyFill(StyleInfo styleInfo)
        {
            return styleInfo.Style.Fill.PatternType.ToOpenXml() == PatternValues.None;
        }

        private static bool ApplyBorder(StyleInfo styleInfo)
        {
            var opBorder = styleInfo.Style.Border;
            return (opBorder.BottomBorder.ToOpenXml() != BorderStyleValues.None
                    || opBorder.DiagonalBorder.ToOpenXml() != BorderStyleValues.None
                    || opBorder.RightBorder.ToOpenXml() != BorderStyleValues.None
                    || opBorder.LeftBorder.ToOpenXml() != BorderStyleValues.None
                    || opBorder.TopBorder.ToOpenXml() != BorderStyleValues.None);
        }

        private static bool ApplyProtection(StyleInfo styleInfo)
        {
            return styleInfo.Style.Protection != null;
        }

        private static CellFormat GetCellFormat(StyleInfo styleInfo)
        {
            var cellFormat = new CellFormat
            {
                NumberFormatId = (UInt32)styleInfo.NumberFormatId,
                FontId = styleInfo.FontId,
                FillId = styleInfo.FillId,
                BorderId = styleInfo.BorderId,
                QuotePrefix = OpenXmlHelper.GetBooleanValue(styleInfo.IncludeQuotePrefix, false),
                ApplyNumberFormat = true,
                ApplyAlignment = true,
                ApplyFill = ApplyFill(styleInfo),
                ApplyBorder = ApplyBorder(styleInfo),
                ApplyProtection = ApplyProtection(styleInfo)
            };
            return cellFormat;
        }

        private static Protection GetProtection(StyleInfo styleInfo)
        {
            return new Protection
            {
                Locked = styleInfo.Style.Protection.Locked,
                Hidden = styleInfo.Style.Protection.Hidden
            };
        }

        /// <summary>
        /// Check if two style are equivalent.
        /// </summary>
        /// <param name="f">Style in the OpenXML format.</param>
        /// <param name="styleInfo">Style in the ClosedXML format.</param>
        /// <param name="compareAlignment">Flag specifying whether or not compare the alignments of two styles.
        /// Styles in x:cellStyleXfs section do not include alignment so we don't have to compare it in this case.
        /// Styles in x:cellXfs section, on the opposite, do include alignments, and we must compare them.
        /// </param>
        /// <returns>True if two formats are equivalent, false otherwise.</returns>
        private static bool CellFormatsAreEqual(CellFormat f, StyleInfo styleInfo, bool compareAlignment)
        {
            return
                f.BorderId != null && styleInfo.BorderId == f.BorderId
                && f.FillId != null && styleInfo.FillId == f.FillId
                && f.FontId != null && styleInfo.FontId == f.FontId
                && f.NumberFormatId != null && styleInfo.NumberFormatId == f.NumberFormatId
                && QuotePrefixesAreEqual(f.QuotePrefix, styleInfo.IncludeQuotePrefix)
                && (f.ApplyFill == null && styleInfo.Style.Fill == XLFillValue.Default ||
                    f.ApplyFill != null && f.ApplyFill == ApplyFill(styleInfo))
                && (f.ApplyBorder == null && styleInfo.Style.Border == XLBorderValue.Default ||
                    f.ApplyBorder != null && f.ApplyBorder == ApplyBorder(styleInfo))
                && (!compareAlignment || AlignmentsAreEqual(f.Alignment, styleInfo.Style.Alignment))
                && ProtectionsAreEqual(f.Protection, styleInfo.Style.Protection)
                ;
        }

        private static bool ProtectionsAreEqual(Protection protection, XLProtectionValue xlProtection)
        {
            var p = XLProtectionValue.Default.Key;
            if (protection != null)
            {
                if (protection.Locked != null)
                    p.Locked = protection.Locked.Value;
                if (protection.Hidden != null)
                    p.Hidden = protection.Hidden.Value;
            }
            return p.Equals(xlProtection.Key);
        }

        private static bool QuotePrefixesAreEqual(BooleanValue quotePrefix, Boolean includeQuotePrefix)
        {
            return OpenXmlHelper.GetBooleanValueAsBool(quotePrefix, false) == includeQuotePrefix;
        }

        private static bool AlignmentsAreEqual(Alignment alignment, XLAlignmentValue xlAlignment)
        {
            if (alignment != null)
            {
                var a = XLAlignmentValue.Default.Key;
                if (alignment.Indent != null)
                    a.Indent = (Int32)alignment.Indent.Value;

                if (alignment.Horizontal != null)
                    a.Horizontal = alignment.Horizontal.Value.ToClosedXml();
                if (alignment.Vertical != null)
                    a.Vertical = alignment.Vertical.Value.ToClosedXml();

                if (alignment.ReadingOrder != null)
                    a.ReadingOrder = alignment.ReadingOrder.Value.ToClosedXml();
                if (alignment.WrapText != null)
                    a.WrapText = alignment.WrapText.Value;
                if (alignment.TextRotation != null)
                    a.TextRotation = OpenXmlHelper.GetClosedXmlTextRotation(alignment);
                if (alignment.ShrinkToFit != null)
                    a.ShrinkToFit = alignment.ShrinkToFit.Value;
                if (alignment.RelativeIndent != null)
                    a.RelativeIndent = alignment.RelativeIndent.Value;
                if (alignment.JustifyLastLine != null)
                    a.JustifyLastLine = alignment.JustifyLastLine.Value;
                return a.Equals(xlAlignment.Key);
            }
            else
            {
                return XLStyle.Default.Value.Alignment.Equals(xlAlignment);
            }
        }

        private static Dictionary<XLBorderValue, BorderInfo> ResolveBorders(WorkbookStylesPart workbookStylesPart,
            Dictionary<XLBorderValue, BorderInfo> sharedBorders)
        {
            if (workbookStylesPart.Stylesheet.Borders == null)
                workbookStylesPart.Stylesheet.Borders = new Borders();

            var allSharedBorders = new Dictionary<XLBorderValue, BorderInfo>();
            foreach (var borderInfo in sharedBorders.Values)
            {
                var borderId = 0;
                var foundOne = false;
                foreach (Border f in workbookStylesPart.Stylesheet.Borders)
                {
                    if (BordersAreEqual(f, borderInfo.Border))
                    {
                        foundOne = true;
                        break;
                    }
                    borderId++;
                }
                if (!foundOne)
                {
                    var border = GetNewBorder(borderInfo);
                    workbookStylesPart.Stylesheet.Borders.AppendChild(border);
                }
                allSharedBorders.Add(borderInfo.Border,
                    new BorderInfo { Border = borderInfo.Border, BorderId = (UInt32)borderId });
            }
            workbookStylesPart.Stylesheet.Borders.Count = (UInt32)workbookStylesPart.Stylesheet.Borders.Count();
            return allSharedBorders;
        }

        private static Border GetNewBorder(BorderInfo borderInfo, Boolean ignoreMod = true)
        {
            var border = new Border();
            if (borderInfo.Border.DiagonalUp != XLBorderValue.Default.DiagonalUp || ignoreMod)
                border.DiagonalUp = borderInfo.Border.DiagonalUp;

            if (borderInfo.Border.DiagonalDown != XLBorderValue.Default.DiagonalDown || ignoreMod)
                border.DiagonalDown = borderInfo.Border.DiagonalDown;

            if (borderInfo.Border.LeftBorder != XLBorderValue.Default.LeftBorder || ignoreMod)
            {
                var leftBorder = new LeftBorder { Style = borderInfo.Border.LeftBorder.ToOpenXml() };
                if (borderInfo.Border.LeftBorderColor != XLBorderValue.Default.LeftBorderColor || ignoreMod)
                {
                    var leftBorderColor = new Color().FromClosedXMLColor<Color>(borderInfo.Border.LeftBorderColor);
                    leftBorder.AppendChild(leftBorderColor);
                }
                border.AppendChild(leftBorder);
            }

            if (borderInfo.Border.RightBorder != XLBorderValue.Default.RightBorder || ignoreMod)
            {
                var rightBorder = new RightBorder { Style = borderInfo.Border.RightBorder.ToOpenXml() };
                if (borderInfo.Border.RightBorderColor != XLBorderValue.Default.RightBorderColor || ignoreMod)
                {
                    var rightBorderColor = new Color().FromClosedXMLColor<Color>(borderInfo.Border.RightBorderColor);
                    rightBorder.AppendChild(rightBorderColor);
                }
                border.AppendChild(rightBorder);
            }

            if (borderInfo.Border.TopBorder != XLBorderValue.Default.TopBorder || ignoreMod)
            {
                var topBorder = new TopBorder { Style = borderInfo.Border.TopBorder.ToOpenXml() };
                if (borderInfo.Border.TopBorderColor != XLBorderValue.Default.TopBorderColor || ignoreMod)
                {
                    var topBorderColor = new Color().FromClosedXMLColor<Color>(borderInfo.Border.TopBorderColor);
                    topBorder.AppendChild(topBorderColor);
                }
                border.AppendChild(topBorder);
            }

            if (borderInfo.Border.BottomBorder != XLBorderValue.Default.BottomBorder || ignoreMod)
            {
                var bottomBorder = new BottomBorder { Style = borderInfo.Border.BottomBorder.ToOpenXml() };
                if (borderInfo.Border.BottomBorderColor != XLBorderValue.Default.BottomBorderColor || ignoreMod)
                {
                    var bottomBorderColor = new Color().FromClosedXMLColor<Color>(borderInfo.Border.BottomBorderColor);
                    bottomBorder.AppendChild(bottomBorderColor);
                }
                border.AppendChild(bottomBorder);
            }

            if (borderInfo.Border.DiagonalBorder != XLBorderValue.Default.DiagonalBorder || ignoreMod)
            {
                var DiagonalBorder = new DiagonalBorder { Style = borderInfo.Border.DiagonalBorder.ToOpenXml() };
                if (borderInfo.Border.DiagonalBorderColor != XLBorderValue.Default.DiagonalBorderColor || ignoreMod)
                    if (borderInfo.Border.DiagonalBorderColor != null)
                    {
                        var DiagonalBorderColor = new Color().FromClosedXMLColor<Color>(borderInfo.Border.DiagonalBorderColor);
                        DiagonalBorder.AppendChild(DiagonalBorderColor);
                    }
                border.AppendChild(DiagonalBorder);
            }

            return border;
        }

        private static bool BordersAreEqual(Border b, XLBorderValue xlBorder)
        {
            var nb = XLBorderValue.Default.Key;
            if (b.DiagonalUp != null)
                nb.DiagonalUp = b.DiagonalUp.Value;

            if (b.DiagonalDown != null)
                nb.DiagonalDown = b.DiagonalDown.Value;

            if (b.DiagonalBorder != null)
            {
                if (b.DiagonalBorder.Style != null)
                    nb.DiagonalBorder = b.DiagonalBorder.Style.Value.ToClosedXml();
                if (b.DiagonalBorder.Color != null)
                    nb.DiagonalBorderColor = b.DiagonalBorder.Color.ToClosedXMLColor().Key;
            }

            if (b.LeftBorder != null)
            {
                if (b.LeftBorder.Style != null)
                    nb.LeftBorder = b.LeftBorder.Style.Value.ToClosedXml();
                if (b.LeftBorder.Color != null)
                    nb.LeftBorderColor = b.LeftBorder.Color.ToClosedXMLColor().Key;
            }

            if (b.RightBorder != null)
            {
                if (b.RightBorder.Style != null)
                    nb.RightBorder = b.RightBorder.Style.Value.ToClosedXml();
                if (b.RightBorder.Color != null)
                    nb.RightBorderColor = b.RightBorder.Color.ToClosedXMLColor().Key;
            }

            if (b.TopBorder != null)
            {
                if (b.TopBorder.Style != null)
                    nb.TopBorder = b.TopBorder.Style.Value.ToClosedXml();
                if (b.TopBorder.Color != null)
                    nb.TopBorderColor = b.TopBorder.Color.ToClosedXMLColor().Key;
            }

            if (b.BottomBorder != null)
            {
                if (b.BottomBorder.Style != null)
                    nb.BottomBorder = b.BottomBorder.Style.Value.ToClosedXml();
                if (b.BottomBorder.Color != null)
                    nb.BottomBorderColor = b.BottomBorder.Color.ToClosedXMLColor().Key;
            }

            return nb.Equals(xlBorder.Key);
        }

        private static Dictionary<XLFillValue, FillInfo> ResolveFills(WorkbookStylesPart workbookStylesPart,
            Dictionary<XLFillValue, FillInfo> sharedFills)
        {
            if (workbookStylesPart.Stylesheet.Fills == null)
                workbookStylesPart.Stylesheet.Fills = new Fills();

            var fills = workbookStylesPart.Stylesheet.Fills;

            // Pattern idx 0 and idx 1 are hardcoded to Excel with values None (0) and Gray125. Excel will ignore
            // values from the file. Every file has have these values inside to keep the first available idx at 2.
            ResolveFillWithPattern(fills, 0, PatternValues.None);
            ResolveFillWithPattern(fills, 1, PatternValues.Gray125);

            var allSharedFills = new Dictionary<XLFillValue, FillInfo>();
            foreach (var fillInfo in sharedFills.Values)
            {
                var fillId = 0;
                var foundOne = false;
                foreach (Fill f in fills)
                {
                    if (FillsAreEqual(f, fillInfo.Fill, fromDifferentialFormat: false))
                    {
                        foundOne = true;
                        break;
                    }
                    fillId++;
                }
                if (!foundOne)
                {
                    var fill = GetNewFill(fillInfo, differentialFillFormat: false);
                    fills.AppendChild(fill);
                }
                allSharedFills.Add(fillInfo.Fill, new FillInfo { Fill = fillInfo.Fill, FillId = (UInt32)fillId });
            }

            fills.Count = (UInt32)fills.Count();
            return allSharedFills;
        }

        private static void ResolveFillWithPattern(Fills fills, Int32 index, PatternValues patternValues)
        {
            var fill = (Fill)fills.ElementAtOrDefault(index);
            if (fill is null)
            {
                fills.InsertAt(new Fill { PatternFill = new PatternFill { PatternType = patternValues } }, index);
                return;
            }

            var fillHasExpectedValue =
                fill.PatternFill?.PatternType?.Value == patternValues &&
                fill.PatternFill.ForegroundColor is null &&
                fill.PatternFill.BackgroundColor is null;

            if (fillHasExpectedValue)
                return;

            fill.PatternFill = new PatternFill { PatternType = patternValues };
        }

        private static Fill GetNewFill(FillInfo fillInfo, Boolean differentialFillFormat, Boolean ignoreMod = true)
        {
            var fill = new Fill();

            var patternFill = new PatternFill();

            patternFill.PatternType = fillInfo.Fill.PatternType.ToOpenXml();

            BackgroundColor backgroundColor;
            ForegroundColor foregroundColor;

            switch (fillInfo.Fill.PatternType)
            {
                case XLFillPatternValues.None:
                    break;

                case XLFillPatternValues.Solid:

                    if (differentialFillFormat)
                    {
                        patternFill.AppendChild(new ForegroundColor { Auto = true });
                        backgroundColor = new BackgroundColor().FromClosedXMLColor<BackgroundColor>(fillInfo.Fill.BackgroundColor, true);
                        if (backgroundColor.HasAttributes)
                            patternFill.AppendChild(backgroundColor);
                    }
                    else
                    {
                        // ClosedXML Background color to be populated into OpenXML fgColor
                        foregroundColor = new ForegroundColor().FromClosedXMLColor<ForegroundColor>(fillInfo.Fill.BackgroundColor);
                        if (foregroundColor.HasAttributes)
                            patternFill.AppendChild(foregroundColor);
                    }
                    break;

                default:

                    foregroundColor = new ForegroundColor().FromClosedXMLColor<ForegroundColor>(fillInfo.Fill.PatternColor);
                    if (foregroundColor.HasAttributes)
                        patternFill.AppendChild(foregroundColor);

                    backgroundColor = new BackgroundColor().FromClosedXMLColor<BackgroundColor>(fillInfo.Fill.BackgroundColor);
                    if (backgroundColor.HasAttributes)
                        patternFill.AppendChild(backgroundColor);

                    break;
            }

            if (patternFill.HasChildren)
                fill.AppendChild(patternFill);

            return fill;
        }

        private static bool FillsAreEqual(Fill f, XLFillValue xlFill, Boolean fromDifferentialFormat)
        {
            var nF = new XLFill(null);

            OpenXmlHelper.LoadFill(f, nF, fromDifferentialFormat);

            return nF.Key.Equals(xlFill.Key);
        }

        private static void ResolveFonts(WorkbookStylesPart workbookStylesPart, SaveContext context)
        {
            if (workbookStylesPart.Stylesheet.Fonts == null)
                workbookStylesPart.Stylesheet.Fonts = new Fonts();

            var newFonts = new Dictionary<XLFontValue, FontInfo>();
            foreach (var fontInfo in context.SharedFonts.Values)
            {
                var fontId = 0;
                var foundOne = false;
                foreach (Font f in workbookStylesPart.Stylesheet.Fonts)
                {
                    if (FontsAreEqual(f, fontInfo.Font))
                    {
                        foundOne = true;
                        break;
                    }
                    fontId++;
                }
                if (!foundOne)
                {
                    var font = GetNewFont(fontInfo);
                    workbookStylesPart.Stylesheet.Fonts.AppendChild(font);
                }
                newFonts.Add(fontInfo.Font, new FontInfo { Font = fontInfo.Font, FontId = (UInt32)fontId });
            }
            context.SharedFonts.Clear();
            foreach (var kp in newFonts)
                context.SharedFonts.Add(kp.Key, kp.Value);

            workbookStylesPart.Stylesheet.Fonts.Count = (UInt32)workbookStylesPart.Stylesheet.Fonts.Count();
        }

        private static Font GetNewFont(FontInfo fontInfo, Boolean ignoreMod = true)
        {
            var font = new Font();
            var bold = (fontInfo.Font.Bold != XLFontValue.Default.Bold || ignoreMod) && fontInfo.Font.Bold ? new Bold() : null;
            var italic = (fontInfo.Font.Italic != XLFontValue.Default.Italic || ignoreMod) && fontInfo.Font.Italic ? new Italic() : null;
            var underline = (fontInfo.Font.Underline != XLFontValue.Default.Underline || ignoreMod) &&
                            fontInfo.Font.Underline != XLFontUnderlineValues.None
                ? new Underline { Val = fontInfo.Font.Underline.ToOpenXml() }
                : null;
            var strike = (fontInfo.Font.Strikethrough != XLFontValue.Default.Strikethrough || ignoreMod) && fontInfo.Font.Strikethrough
                ? new Strike()
                : null;
            var verticalAlignment = fontInfo.Font.VerticalAlignment != XLFontValue.Default.VerticalAlignment || ignoreMod
                ? new VerticalTextAlignment { Val = fontInfo.Font.VerticalAlignment.ToOpenXml() }
                : null;

            var shadow = (fontInfo.Font.Shadow != XLFontValue.Default.Shadow || ignoreMod) && fontInfo.Font.Shadow ? new Shadow() : null;
            var fontSize = fontInfo.Font.FontSize != XLFontValue.Default.FontSize || ignoreMod
                ? new FontSize { Val = fontInfo.Font.FontSize }
                : null;
            var color = fontInfo.Font.FontColor != XLFontValue.Default.FontColor || ignoreMod ? new Color().FromClosedXMLColor<Color>(fontInfo.Font.FontColor) : null;

            var fontName = fontInfo.Font.FontName != XLFontValue.Default.FontName || ignoreMod
                ? new FontName { Val = fontInfo.Font.FontName }
                : null;
            var fontFamilyNumbering = fontInfo.Font.FontFamilyNumbering != XLFontValue.Default.FontFamilyNumbering || ignoreMod
                ? new FontFamilyNumbering { Val = (Int32)fontInfo.Font.FontFamilyNumbering }
                : null;

            var fontCharSet = (fontInfo.Font.FontCharSet != XLFontValue.Default.FontCharSet || ignoreMod) && fontInfo.Font.FontCharSet != XLFontCharSet.Default
                ? new FontCharSet { Val = (Int32)fontInfo.Font.FontCharSet }
                : null;

            var fontScheme = (fontInfo.Font.FontScheme != XLFontValue.Default.FontScheme || ignoreMod) && fontInfo.Font.FontScheme != XLFontScheme.None
                ? new DocumentFormat.OpenXml.Spreadsheet.FontScheme { Val = fontInfo.Font.FontScheme.ToOpenXmlEnum() }
                : null;

            if (bold != null)
                font.AppendChild(bold);
            if (italic != null)
                font.AppendChild(italic);
            if (underline != null)
                font.AppendChild(underline);
            if (strike != null)
                font.AppendChild(strike);
            if (verticalAlignment != null)
                font.AppendChild(verticalAlignment);
            if (shadow != null)
                font.AppendChild(shadow);
            if (fontSize != null)
                font.AppendChild(fontSize);
            if (color != null)
                font.AppendChild(color);
            if (fontName != null)
                font.AppendChild(fontName);
            if (fontFamilyNumbering != null)
                font.AppendChild(fontFamilyNumbering);
            if (fontCharSet != null)
                font.AppendChild(fontCharSet);
            if (fontScheme != null)
                font.AppendChild(fontScheme);

            return font;
        }

        private static bool FontsAreEqual(Font f, XLFontValue xlFont)
        {
            var nf = XLFontValue.Default.Key;
            nf.Bold = f.Bold != null;
            nf.Italic = f.Italic != null;

            if (f.Underline != null)
            {
                nf.Underline = f.Underline.Val != null
                    ? f.Underline.Val.Value.ToClosedXml()
                    : XLFontUnderlineValues.Single;
            }
            nf.Strikethrough = f.Strike != null;
            if (f.VerticalTextAlignment != null)
            {
                nf.VerticalAlignment = f.VerticalTextAlignment.Val != null
                    ? f.VerticalTextAlignment.Val.Value.ToClosedXml()
                    : XLFontVerticalTextAlignmentValues.Baseline;
            }
            nf.Shadow = f.Shadow != null;
            if (f.FontSize != null)
                nf.FontSize = f.FontSize.Val;
            if (f.Color != null)
                nf.FontColor = f.Color.ToClosedXMLColor().Key;
            if (f.FontName != null)
                nf.FontName = f.FontName.Val;
            if (f.FontFamilyNumbering != null)
                nf.FontFamilyNumbering = (XLFontFamilyNumberingValues)f.FontFamilyNumbering.Val.Value;
            if (f.FontScheme?.Val != null)
                nf.FontScheme = f.FontScheme.Val.Value.ToClosedXml();

            return nf.Equals(xlFont.Key);
        }

        private static Dictionary<XLNumberFormatValue, NumberFormatInfo> ResolveNumberFormats(
            WorkbookStylesPart workbookStylesPart,
            HashSet<XLNumberFormatValue> customNumberFormats,
            UInt32 defaultFormatId)
        {
            if (workbookStylesPart.Stylesheet.NumberingFormats == null)
            {
                workbookStylesPart.Stylesheet.NumberingFormats = new NumberingFormats();
                workbookStylesPart.Stylesheet.NumberingFormats.AppendChild(new NumberingFormat()
                {
                    NumberFormatId = 0,
                    FormatCode = ""
                });
            }

            var allSharedNumberFormats = new Dictionary<XLNumberFormatValue, NumberFormatInfo>();
            var partNumberingFormats = workbookStylesPart.Stylesheet.NumberingFormats;

            // number format ids in the part can have holes in the sequence and first id can be greater than last built-in style id.
            // In some cases, there are also existing number formats with id below last built-in style id.
            var availableNumberFormatId = partNumberingFormats.Any()
                ? Math.Max(partNumberingFormats.Cast<NumberingFormat>().Max(nf => nf.NumberFormatId!.Value) + 1, XLConstants.NumberOfBuiltInStyles)
                : XLConstants.NumberOfBuiltInStyles; // 0-based

            // Merge custom formats used in the workbook that are not already present in the part to the part and assign ids
            foreach (var customNumberFormat in customNumberFormats.Where(nf => nf.NumberFormatId != defaultFormatId))
            {
                NumberingFormat partNumberFormat = null;
                foreach (var nf in workbookStylesPart.Stylesheet.NumberingFormats.Cast<NumberingFormat>())
                {
                    if (CustomNumberFormatsAreEqual(nf, customNumberFormat))
                    {
                        partNumberFormat = nf;
                        break;
                    }
                }
                if (partNumberFormat is null)
                {
                    partNumberFormat = new NumberingFormat
                    {
                        NumberFormatId = (UInt32)availableNumberFormatId++,
                        FormatCode = customNumberFormat.Format
                    };
                    workbookStylesPart.Stylesheet.NumberingFormats.AppendChild(partNumberFormat);
                }
                allSharedNumberFormats.Add(customNumberFormat,
                    new NumberFormatInfo
                    {
                        NumberFormat = customNumberFormat,
                        NumberFormatId = (Int32)partNumberFormat.NumberFormatId!.Value
                    });
            }
            workbookStylesPart.Stylesheet.NumberingFormats.Count =
                (UInt32)workbookStylesPart.Stylesheet.NumberingFormats.Count();
            return allSharedNumberFormats;
        }

        private static bool CustomNumberFormatsAreEqual(NumberingFormat nf, XLNumberFormatValue xlNumberFormat)
        {
            if (nf.FormatCode != null && !String.IsNullOrWhiteSpace(nf.FormatCode.Value))
                return string.Equals(xlNumberFormat?.Format, nf.FormatCode.Value);

            return false;
        }


    }
}
