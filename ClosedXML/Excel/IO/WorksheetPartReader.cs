using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.Collections.Generic;
using System;
using System.Globalization;
using System.Linq;
using ClosedXML.Utils;
using ClosedXML.Extensions;
using ClosedXML.IO;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using System.Diagnostics;
using ClosedXML.Parser;
using static ClosedXML.Excel.XLPredefinedFormat.DateTime;

namespace ClosedXML.Excel.IO;

#nullable disable

internal class WorksheetPartReader
{
    private static readonly string[] DateCellFormats =
    {
        "yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fff", // Format accepted by OpenXML SDK
        "yyyy-MM-ddTHH:mm", "yyyy-MM-dd" // Formats accepted by Excel.
    };

    private readonly Dictionary<UInt32, String> _sharedFormulasR1C1 = new();
    private Int32 _lastRow;
    private Int32 _lastColumnNumber;

    internal void LoadWorksheet(XLWorksheet ws, Stylesheet s, Fills fills, Borders borders, NumberingFormats numberingFormats, WorksheetPart worksheetPart, SharedStringItem[] sharedStrings, Dictionary<int, DifferentialFormat> differentialFormats, LoadContext context)
    {
        ApplyStyle(ws, 0, s, fills, borders, numberingFormats, ws.Workbook.Styles);

        var styleList = new Dictionary<int, IXLStyle>();// {{0, ws.Style}};
        PageSetupProperties pageSetupProperties = null;

        _lastRow = 0;

        using (var reader = new OpenXmlPartReader(worksheetPart))
        {
            Type[] ignoredElements = new Type[]
            {
                    typeof(CustomSheetViews) // Custom sheet views contain its own auto filter data, and more, which should be ignored for now
            };

            while (reader.Read())
            {
                while (ignoredElements.Contains(reader.ElementType))
                    reader.ReadNextSibling();

                if (reader.ElementType == typeof(SheetFormatProperties))
                {
                    var sheetFormatProperties = (SheetFormatProperties)reader.LoadCurrentElement();
                    if (sheetFormatProperties != null)
                    {
                        if (sheetFormatProperties.DefaultRowHeight != null)
                            ws.RowHeight = sheetFormatProperties.DefaultRowHeight;

                        ws.RowHeightChanged = (sheetFormatProperties.CustomHeight != null &&
                                               sheetFormatProperties.CustomHeight.Value);

                        if (sheetFormatProperties.DefaultColumnWidth != null)
                            ws.ColumnWidth = XLHelper.ConvertWidthToNoC(sheetFormatProperties.DefaultColumnWidth.Value, ws.Style.Font, ws.Workbook);
                        else if (sheetFormatProperties.BaseColumnWidth != null)
                            ws.ColumnWidth = XLHelper.CalculateColumnWidth(sheetFormatProperties.BaseColumnWidth.Value, ws.Style.Font, ws.Workbook);
                    }
                }
                else if (reader.ElementType == typeof(SheetViews))
                    LoadSheetViews((SheetViews)reader.LoadCurrentElement(), ws);
                else if (reader.ElementType == typeof(MergeCells))
                {
                    var mergedCells = (MergeCells)reader.LoadCurrentElement();
                    if (mergedCells != null)
                    {
                        foreach (MergeCell mergeCell in mergedCells.Elements<MergeCell>())
                            ws.Range(mergeCell.Reference).Merge(false);
                    }
                }
                else if (reader.ElementType == typeof(Columns))
                    LoadColumns(s, numberingFormats, fills, borders, ws,
                        (Columns)reader.LoadCurrentElement());
                else if (reader.ElementType == typeof(Row))
                {
                    LoadRow(s, numberingFormats, fills, borders, ws, sharedStrings, styleList, reader);
                }
                else if (reader.ElementType == typeof(AutoFilter))
                    AutoFilterReader.LoadAutoFilter((AutoFilter)reader.LoadCurrentElement(), ws);
                else if (reader.ElementType == typeof(SheetProtection))
                    LoadSheetProtection((SheetProtection)reader.LoadCurrentElement(), ws);
                else if (reader.ElementType == typeof(DataValidations))
                    LoadDataValidations((DataValidations)reader.LoadCurrentElement(), ws);
                else if (reader.ElementType == typeof(ConditionalFormatting))
                    LoadConditionalFormatting((ConditionalFormatting)reader.LoadCurrentElement(), ws, differentialFormats, context);
                else if (reader.ElementType == typeof(Hyperlinks))
                    LoadHyperlinks((Hyperlinks)reader.LoadCurrentElement(), worksheetPart, ws);
                else if (reader.ElementType == typeof(PrintOptions))
                    LoadPrintOptions((PrintOptions)reader.LoadCurrentElement(), ws);
                else if (reader.ElementType == typeof(PageMargins))
                    LoadPageMargins((PageMargins)reader.LoadCurrentElement(), ws);
                else if (reader.ElementType == typeof(PageSetup))
                    LoadPageSetup((PageSetup)reader.LoadCurrentElement(), ws, pageSetupProperties);
                else if (reader.ElementType == typeof(HeaderFooter))
                    LoadHeaderFooter((HeaderFooter)reader.LoadCurrentElement(), ws);
                else if (reader.ElementType == typeof(SheetProperties))
                    LoadSheetProperties((SheetProperties)reader.LoadCurrentElement(), ws, out pageSetupProperties);
                else if (reader.ElementType == typeof(RowBreaks))
                    LoadRowBreaks((RowBreaks)reader.LoadCurrentElement(), ws);
                else if (reader.ElementType == typeof(ColumnBreaks))
                    LoadColumnBreaks((ColumnBreaks)reader.LoadCurrentElement(), ws);
                else if (reader.ElementType == typeof(WorksheetExtensionList))
                    LoadExtensions((WorksheetExtensionList)reader.LoadCurrentElement(), ws);
                else if (reader.ElementType == typeof(LegacyDrawing))
                    ws.LegacyDrawingId = (reader.LoadCurrentElement() as LegacyDrawing).Id.Value;
            }
            reader.Close();
        }
    }

    private static void LoadSheetProperties(SheetProperties sheetProperty, XLWorksheet ws, out PageSetupProperties pageSetupProperties)
    {
        pageSetupProperties = null;
        if (sheetProperty == null) return;

        if (sheetProperty.TabColor != null)
            ws.TabColor = sheetProperty.TabColor.ToClosedXMLColor();

        if (sheetProperty.OutlineProperties != null)
        {
            if (sheetProperty.OutlineProperties.SummaryBelow != null)
            {
                ws.Outline.SummaryVLocation = sheetProperty.OutlineProperties.SummaryBelow
                    ? XLOutlineSummaryVLocation.Bottom
                    : XLOutlineSummaryVLocation.Top;
            }

            if (sheetProperty.OutlineProperties.SummaryRight != null)
            {
                ws.Outline.SummaryHLocation = sheetProperty.OutlineProperties.SummaryRight
                    ? XLOutlineSummaryHLocation.Right
                    : XLOutlineSummaryHLocation.Left;
            }
        }

        if (sheetProperty.PageSetupProperties != null)
            pageSetupProperties = sheetProperty.PageSetupProperties;
    }

    private static void LoadColumns(Stylesheet s, NumberingFormats numberingFormats, Fills fills, Borders borders,
                             XLWorksheet ws, Columns columns)
    {
        if (columns == null) return;

        var wsDefaultColumn =
            columns.Elements<Column>().FirstOrDefault(c => c.Max == XLHelper.MaxColumnNumber);

        if (wsDefaultColumn != null && wsDefaultColumn.Width != null)
            ws.ColumnWidth = wsDefaultColumn.Width - XLConstants.ColumnWidthOffset;

        Int32 styleIndexDefault = wsDefaultColumn != null && wsDefaultColumn.Style != null
                                      ? Int32.Parse(wsDefaultColumn.Style.InnerText)
                                      : -1;
        if (styleIndexDefault >= 0)
            ApplyStyle(ws, styleIndexDefault, s, fills, borders, numberingFormats, ws.Workbook.Styles);

        foreach (Column col in columns.Elements<Column>())
        {
            //IXLStylized toApply;
            if (col.Max == XLHelper.MaxColumnNumber) continue;

            var xlColumns = (XLColumns)ws.Columns(col.Min, col.Max);
            if (col.Width != null)
            {
                Double width = col.Width - XLConstants.ColumnWidthOffset;
                //if (width < 0) width = 0;
                xlColumns.Width = width;
            }
            else
                xlColumns.Width = ws.ColumnWidth;

            if (col.Hidden != null && col.Hidden)
                xlColumns.Hide();

            if (col.Collapsed != null && col.Collapsed)
                xlColumns.CollapseOnly();

            if (col.OutlineLevel != null)
            {
                var outlineLevel = col.OutlineLevel;
                xlColumns.ForEach(c => c.OutlineLevel = outlineLevel);
            }

            Int32 styleIndex = col.Style != null ? Int32.Parse(col.Style.InnerText) : -1;
            if (styleIndex >= 0)
            {
                ApplyStyle(xlColumns, styleIndex, s, fills, borders, numberingFormats, ws.Workbook.Styles);
            }
            else
            {
                xlColumns.Style = ws.Style;
            }
        }
    }


    private void LoadRow(Stylesheet s, NumberingFormats numberingFormats, Fills fills, Borders borders,
                          XLWorksheet ws, SharedStringItem[] sharedStrings,
                          Dictionary<Int32, IXLStyle> styleList,
                          OpenXmlPartReader reader)
    {
        Debug.Assert(reader.LocalName == "row");

        var attributes = reader.Attributes;
        var rowIndexAttr = attributes.GetAttribute("r");
        var rowIndex = string.IsNullOrEmpty(rowIndexAttr) ? ++_lastRow : int.Parse(rowIndexAttr);

        var xlRow = ws.Row(rowIndex, false);

        var height = attributes.GetDoubleAttribute("ht");
        if (height is not null)
        {
            xlRow.Height = height.Value;
        }
        else
        {
            xlRow.Loading = true;
            xlRow.Height = ws.RowHeight;
            xlRow.Loading = false;
        }

        var dyDescent = attributes.GetDoubleAttribute("dyDescent", OpenXmlConst.X14Ac2009SsNs);
        if (dyDescent is not null)
            xlRow.DyDescent = dyDescent.Value;

        var hidden = attributes.GetBoolAttribute("hidden", false);
        if (hidden)
            xlRow.Hide();

        var collapsed = attributes.GetBoolAttribute("collapsed", false);
        if (collapsed)
            xlRow.Collapsed = true;

        var outlineLevel = attributes.GetIntAttribute("outlineLevel");
        if (outlineLevel is not null && outlineLevel.Value > 0)
            xlRow.OutlineLevel = outlineLevel.Value;

        var showPhonetic = attributes.GetBoolAttribute("ph", false);
        if (showPhonetic)
            xlRow.ShowPhonetic = true;

        var customFormat = attributes.GetBoolAttribute("customFormat", false);
        if (customFormat)
        {
            var styleIndex = attributes.GetIntAttribute("s");
            if (styleIndex is not null)
            {
                ApplyStyle(xlRow, styleIndex.Value, s, fills, borders, numberingFormats, ws.Workbook.Styles);
            }
            else
            {
                xlRow.Style = ws.Style;
            }
        }

        _lastColumnNumber = 0;

        // Move from the start element of 'row' forward. We can get cell, extList or end of row.
        reader.MoveAhead();

        while (reader.IsStartElement("c"))
        {
            LoadCell(sharedStrings, s, numberingFormats, fills, borders, ws, styleList,
                reader, rowIndex);

            // Move from end element of 'cell' either to next cell, extList start or end of row.
            reader.MoveAhead();
        }

        // In theory, row can also contain extList, just skip them.
        while (reader.IsStartElement("extLst"))
            reader.Skip();
    }

    private void LoadCell(SharedStringItem[] sharedStrings, Stylesheet s, NumberingFormats numberingFormats,
                          Fills fills, Borders borders,
                          XLWorksheet ws, Dictionary<Int32, IXLStyle> styleList, OpenXmlPartReader reader, Int32 rowIndex)
    {
        Debug.Assert(reader.LocalName == "c" && reader.IsStartElement);

        var attributes = reader.Attributes;

        var styleIndex = attributes.GetIntAttribute("s") ?? 0;

        var cellAddress = attributes.GetCellRefAttribute("r") ?? new XLSheetPoint(rowIndex, _lastColumnNumber + 1);
        _lastColumnNumber = cellAddress.Column;

        var dataType = attributes.GetAttribute("t") switch
        {
            "b" => CellValues.Boolean,
            "n" => CellValues.Number,
            "e" => CellValues.Error,
            "s" => CellValues.SharedString,
            "str" => CellValues.String,
            "inlineStr" => CellValues.InlineString,
            "d" => CellValues.Date,
            null => CellValues.Number,
            _ => throw new FormatException($"Unknown cell type.")
        };

        var xlCell = ws.Cell(cellAddress.Row, cellAddress.Column);

        if (styleList.TryGetValue(styleIndex, out IXLStyle style))
        {
            xlCell.InnerStyle = style;
        }
        else
        {
            ApplyStyle(xlCell, styleIndex, s, fills, borders, numberingFormats, ws.Workbook.Styles);
        }

        var showPhonetic = attributes.GetBoolAttribute("ph", false);
        if (showPhonetic)
            xlCell.ShowPhonetic = true;

        var cellMetaIndex = attributes.GetUintAttribute("cm");
        if (cellMetaIndex is not null)
            xlCell.CellMetaIndex = cellMetaIndex.Value;

        var valueMetaIndex = attributes.GetUintAttribute("vm");
        if (valueMetaIndex is not null)
            xlCell.ValueMetaIndex = valueMetaIndex.Value;

        // Move from cell start element onwards.
        reader.MoveAhead();

        var cellHasFormula = reader.IsStartElement("f");
        XLCellFormula formula = null;
        if (cellHasFormula)
        {
            formula = SetCellFormula(ws, cellAddress, reader);

            // Move from end of 'f' element.
            reader.MoveAhead();
        }

        // Unified code to load value. Value can be empty and only type specified (e.g. when formula doesn't save values)
        // String type is only for formulas, while shared string/inline string/date is only for pure cell values.
        var cellHasValue = reader.IsStartElement("v");
        if (cellHasValue)
        {
            SetCellValue(dataType, reader.GetText(), xlCell, sharedStrings);

            // Skips all nodes of the 'v' element (has no child nodes) and moves to the first element after.
            reader.Skip();
        }
        else
        {
            // A string cell must contain at least empty string.
            if (dataType.Equals(CellValues.SharedString) || dataType.Equals(CellValues.String))
                xlCell.SetOnlyValue(string.Empty);
        }

        // If the cell doesn't contain value, we should invalidate it, otherwise rely on the stored value.
        // The value is likely more reliable. It should be set when cellFormula.CalculateCell is set or
        // when value is missing. Formula can be null in some cases, e.g. slave cells of array formula.
        if (formula is not null && !cellHasValue)
        {
            formula.IsDirty = true;
        }

        // Inline text is dealt separately, because it is in a separate element.
        var cellHasInlineString = reader.IsStartElement("is");
        if (cellHasInlineString)
        {
            if (dataType == CellValues.InlineString)
            {
                xlCell.ShareString = false;
                var inlineString = (RstType)reader.LoadCurrentElement();
                if (inlineString is not null)
                {
                    if (inlineString.Text is not null)
                        xlCell.SetOnlyValue(inlineString.Text.Text.FixNewLines());
                    else
                        SetCellText(xlCell, inlineString);
                }
                else
                {
                    xlCell.SetOnlyValue(String.Empty);
                }

                // Move from end 'is' element to the end of a 'c' element.
                reader.MoveAhead();
            }
            else
            {
                // Move to the first node after end of 'is' element, which should be end of cell.
                reader.Skip();
            }
        }

        if (ws.Workbook.Use1904DateSystem && xlCell.DataType == XLDataType.DateTime)
        {
            // Internally ClosedXML stores cells as standard 1900-based style
            // so if a workbook is in 1904-format, we do that adjustment here and when saving.
            xlCell.SetOnlyValue(xlCell.GetDateTime().AddDays(1462));
        }

        if (!styleList.ContainsKey(styleIndex))
            styleList.Add(styleIndex, xlCell.Style);
    }

    private XLCellFormula SetCellFormula(XLWorksheet ws, XLSheetPoint cellAddress, OpenXmlPartReader reader)
    {
        var attributes = reader.Attributes;
        var formulaSlice = ws.Internals.CellsCollection.FormulaSlice;
        var valueSlice = ws.Internals.CellsCollection.ValueSlice;

        // bx attribute of cell formula is not ever used, per MS-OI29500 2.1.620
        var formulaText = reader.GetText();
        var formulaType = attributes.GetAttribute("t") switch
        {
            "normal" => CellFormulaValues.Normal,
            "array" => CellFormulaValues.Array,
            "dataTable" => CellFormulaValues.DataTable,
            "shared" => CellFormulaValues.Shared,
            null => CellFormulaValues.Normal,
            _ => throw new NotSupportedException("Unknown formula type.")
        };

        // Always set shareString flag to `false`, because the text result of
        // formula is stored directly in the sheet, not shared string table.
        XLCellFormula formula = null;
        if (formulaType == CellFormulaValues.Normal)
        {
            formula = XLCellFormula.NormalA1(formulaText);
            formulaSlice.Set(cellAddress, formula);
            valueSlice.SetShareString(cellAddress, false);
        }
        else if (formulaType == CellFormulaValues.Array && attributes.GetRefAttribute("ref") is { } arrayArea) // Child cells of an array may have array type, but not ref, that is reserved for master cell
        {
            var aca = attributes.GetBoolAttribute("aca", false);

            // Because cells are read from top-to-bottom, from left-to-right, none of child cells have
            // a formula yet. Also, Excel doesn't allow change of array data, only through parent formula.
            formula = XLCellFormula.Array(formulaText, arrayArea, aca);
            formulaSlice.SetArray(arrayArea, formula);

            for (var col = arrayArea.FirstPoint.Column; col <= arrayArea.LastPoint.Column; ++col)
            {
                for (var row = arrayArea.FirstPoint.Row; row <= arrayArea.LastPoint.Row; ++row)
                {
                    valueSlice.SetShareString(cellAddress, false);
                }
            }
        }
        else if (formulaType == CellFormulaValues.Shared && attributes.GetUintAttribute("si") is { } sharedIndex)
        {
            // Shared formulas are rather limited in use and parsing, even by Excel
            // https://stackoverflow.com/questions/54654993. Therefore we accept them,
            // but don't output them. Shared formula is created, when user in Excel
            // takes a supported formula and drags it to more cells.
            if (!_sharedFormulasR1C1.TryGetValue(sharedIndex, out var sharedR1C1Formula))
            {
                // Spec: The first formula in a group of shared formulas is saved
                // in the f element. This is considered the 'master' formula cell.
                formula = XLCellFormula.NormalA1(formulaText);
                formulaSlice.Set(cellAddress, formula);

                // The key reason why Excel hates shared formulas is likely relative addressing and the messy situation it creates
                var formulaR1C1 = FormulaConverter.ToR1C1(formulaText, cellAddress.Row, cellAddress.Column);
                _sharedFormulasR1C1.Add(sharedIndex, formulaR1C1);
            }
            else
            {
                // Spec: The formula expression for a cell that is specified to be part of a shared formula
                // (and is not the master) shall be ignored, and the master formula shall override.
                var sharedFormulaA1 = FormulaConverter.ToA1(sharedR1C1Formula, cellAddress.Row, cellAddress.Column);
                formula = XLCellFormula.NormalA1(sharedFormulaA1);
                formulaSlice.Set(cellAddress, formula);
            }

            valueSlice.SetShareString(cellAddress, false);
        }
        else if (formulaType == CellFormulaValues.DataTable && attributes.GetRefAttribute("ref") is { } dataTableArea)
        {
            var is2D = attributes.GetBoolAttribute("dt2D", false);
            var input1Deleted = attributes.GetBoolAttribute("del1", false);
            var input1 = attributes.GetCellRefAttribute("r1") ?? throw PartStructureException.MissingAttribute("r1");
            if (is2D)
            {
                // Input 2 is only used for 2D tables
                var input2Deleted = attributes.GetBoolAttribute("del2", false);
                var input2 = attributes.GetCellRefAttribute("r2") ?? throw PartStructureException.MissingAttribute("r2");
                formula = XLCellFormula.DataTable2D(dataTableArea, input1, input1Deleted, input2, input2Deleted);
                formulaSlice.Set(cellAddress, formula);
            }
            else
            {
                var isRowDataTable = attributes.GetBoolAttribute("dtr", false);
                formula = XLCellFormula.DataTable1D(dataTableArea, input1, input1Deleted, isRowDataTable);
                formulaSlice.Set(cellAddress, formula);
            }

            valueSlice.SetShareString(cellAddress, false);
        }

        // Go from start of 'f' element to the end of 'f' element.
        reader.MoveAhead();

        return formula;
    }

    private void SetCellValue(CellValues dataType, string cellValue, XLCell xlCell, SharedStringItem[] sharedStrings)
    {
        if (dataType == CellValues.Number)
        {
            // XLCell is by default blank, so no need to set it.
            if (cellValue is not null && double.TryParse(cellValue, XLHelper.NumberStyle, XLHelper.ParseCulture, out var number))
            {
                var numberDataType = GetNumberDataType(xlCell.StyleValue.NumberFormat);
                var cellNumber = numberDataType switch
                {
                    XLDataType.DateTime => XLCellValue.FromSerialDateTime(number),
                    XLDataType.TimeSpan => XLCellValue.FromSerialTimeSpan(number),
                    _ => number // Normal number
                };
                xlCell.SetOnlyValue(cellNumber);
            }
        }
        else if (dataType == CellValues.SharedString)
        {
            if (cellValue is not null
                && Int32.TryParse(cellValue, XLHelper.NumberStyle, XLHelper.ParseCulture, out Int32 sharedStringId)
                && sharedStringId >= 0 && sharedStringId < sharedStrings.Length)
            {
                var sharedString = sharedStrings[sharedStringId];

                SetCellText(xlCell, sharedString);
            }
            else
                xlCell.SetOnlyValue(String.Empty);
        }
        else if (dataType == CellValues.String) // A plain string that is a result of a formula calculation
        {
            xlCell.SetOnlyValue(cellValue ?? String.Empty);
        }
        else if (dataType == CellValues.Boolean)
        {
            if (cellValue is not null)
            {
                var isTrue = string.Equals(cellValue, "1", StringComparison.Ordinal) ||
                             string.Equals(cellValue, "TRUE", StringComparison.OrdinalIgnoreCase);
                xlCell.SetOnlyValue(isTrue);
            }
        }
        else if (dataType == CellValues.Error)
        {
            if (cellValue is not null && XLErrorParser.TryParseError(cellValue, out var error))
                xlCell.SetOnlyValue(error);
        }
        else if (dataType == CellValues.Date)
        {
            // Technically, cell can contain date as ISO8601 string, but not rarely used due
            // to inconsistencies between ISO and serial date time representation.
            if (cellValue is not null)
            {
                var date = DateTime.ParseExact(cellValue, DateCellFormats,
                    XLHelper.ParseCulture,
                    DateTimeStyles.AllowLeadingWhite | DateTimeStyles.AllowTrailingWhite);
                xlCell.SetOnlyValue(date);
            }
        }
    }

    /// <summary>
    /// Parses the cell value for normal or rich text
    /// Input element should either be a shared string or inline string
    /// </summary>
    /// <param name="xlCell">The cell.</param>
    /// <param name="element">The element (either a shared string or inline string)</param>
    private void SetCellText(XLCell xlCell, RstType element)
    {
        var runs = element.Elements<Run>();
        var hasRuns = false;
        foreach (Run run in runs)
        {
            hasRuns = true;
            var runProperties = run.RunProperties;
            String text = run.Text.InnerText.FixNewLines();

            if (runProperties == null)
                xlCell.GetRichText().AddText(text, xlCell.Style.Font);
            else
            {
                var rt = xlCell.GetRichText().AddText(text);
                var fontScheme = runProperties.Elements<FontScheme>().FirstOrDefault();
                if (fontScheme != null && fontScheme.Val is not null)
                    rt.SetFontScheme(fontScheme.Val.Value.ToClosedXml());

                OpenXmlHelper.LoadFont(runProperties, rt);
            }
        }

        if (!hasRuns)
            xlCell.SetOnlyValue(XStringConvert.Decode(element.Text?.InnerText) ?? string.Empty);

        // Load phonetic properties
        var phoneticProperties = element.Elements<PhoneticProperties>();
        var pp = phoneticProperties.FirstOrDefault();
        if (pp != null)
        {
            if (pp.Alignment != null)
                xlCell.GetRichText().Phonetics.Alignment = pp.Alignment.Value.ToClosedXml();
            if (pp.Type != null)
                xlCell.GetRichText().Phonetics.Type = pp.Type.Value.ToClosedXml();

            OpenXmlHelper.LoadFont(pp, xlCell.GetRichText().Phonetics);
        }

        // Load phonetic runs
        var phoneticRuns = element.Elements<PhoneticRun>();
        foreach (PhoneticRun pr in phoneticRuns)
        {
            xlCell.GetRichText().Phonetics.Add(pr.Text.InnerText.FixNewLines(), (Int32)pr.BaseTextStartIndex.Value,
                                          (Int32)pr.EndingBaseIndex.Value);
        }
    }

    private static XLDataType GetNumberDataType(XLNumberFormatValue numberFormat)
    {
        var numberFormatId = (XLPredefinedFormat.DateTime)numberFormat.NumberFormatId;
        var isTimeOnlyFormat = numberFormatId is
            Hour12MinutesAmPm or
            Hour12MinutesSecondsAmPm or
            Hour24Minutes or
            Hour24MinutesSeconds or
            MinutesSeconds or
            Hour12MinutesSeconds or
            MinutesSecondsMillis1;

        if (isTimeOnlyFormat)
            return XLDataType.TimeSpan;

        var isDateTimeFormat = numberFormatId is
                DayMonthYear4WithSlashes or
                DayMonthAbbrYear2WithDashes or
                DayMonthAbbrWithDash or
                MonthDayYear4WithDashesHour24Minutes;

        if (isDateTimeFormat)
            return XLDataType.DateTime;

        if (!String.IsNullOrWhiteSpace(numberFormat.Format))
        {
            var dataType = GetDataTypeFromFormat(numberFormat.Format);
            return dataType ?? XLDataType.Number;
        }

        return XLDataType.Number;
    }

    private static XLDataType? GetDataTypeFromFormat(String format)
    {
        int length = format.Length;
        String f = format.ToLower();
        for (Int32 i = 0; i < length; i++)
        {
            Char c = f[i];
            if (c == '"')
                i = f.IndexOf('"', i + 1);
            else if (c == '[')
            {
                // #1742 We need to skip locale prefixes in DateTime formats [...]
                i = f.IndexOf(']', i + 1);
                if (i == -1)
                    return null;
            }
            else if (c == '0' || c == '#' || c == '?')
                return XLDataType.Number;
            else if (c == 'y' || c == 'd')
                return XLDataType.DateTime;
            else if (c == 'h' || c == 's')
                return XLDataType.TimeSpan;
            else if (c == 'm')
            {
                // Excel treats "m" immediately after "hh" or "h" or immediately before "ss" or "s" as minutes, otherwise as a month value
                // We can ignore the "hh" or "h" prefixes as these would have been detected by the preceding condition above.
                // So we just need to make sure any 'm' is followed immediately by "ss" or "s" (excluding placeholders) to detect a timespan value
                for (Int32 j = i + 1; j < length; j++)
                {
                    if (f[j] == 'm')
                        continue;
                    else if (f[j] == 's')
                        return XLDataType.TimeSpan;
                    else if ((f[j] >= 'a' && f[j] <= 'z') || (f[j] >= '0' && f[j] <= '9'))
                        return XLDataType.DateTime;
                }
                return XLDataType.DateTime;
            }
        }
        return null;
    }

    private static void LoadSheetViews(SheetViews sheetViews, XLWorksheet ws)
    {
        if (sheetViews == null) return;

        var sheetView = sheetViews.Elements<SheetView>().FirstOrDefault();

        if (sheetView == null) return;

        if (sheetView.RightToLeft != null) ws.RightToLeft = sheetView.RightToLeft.Value;
        if (sheetView.ShowFormulas != null) ws.ShowFormulas = sheetView.ShowFormulas.Value;
        if (sheetView.ShowGridLines != null) ws.ShowGridLines = sheetView.ShowGridLines.Value;
        if (sheetView.ShowOutlineSymbols != null)
            ws.ShowOutlineSymbols = sheetView.ShowOutlineSymbols.Value;
        if (sheetView.ShowRowColHeaders != null) ws.ShowRowColHeaders = sheetView.ShowRowColHeaders.Value;
        if (sheetView.ShowRuler != null) ws.ShowRuler = sheetView.ShowRuler.Value;
        if (sheetView.ShowWhiteSpace != null) ws.ShowWhiteSpace = sheetView.ShowWhiteSpace.Value;
        if (sheetView.ShowZeros != null) ws.ShowZeros = sheetView.ShowZeros.Value;
        if (sheetView.TabSelected != null) ws.TabSelected = sheetView.TabSelected.Value;

        var selection = sheetView.Elements<Selection>().FirstOrDefault();
        if (selection != null)
        {
            if (selection.SequenceOfReferences != null)
                ws.Ranges(selection.SequenceOfReferences.InnerText.Replace(" ", ",")).Select();

            if (selection.ActiveCell != null)
                ws.Cell(selection.ActiveCell).SetActive();
        }

        if (sheetView.ZoomScale != null)
            ws.SheetView.ZoomScale = (int)UInt32Value.ToUInt32(sheetView.ZoomScale);
        if (sheetView.ZoomScaleNormal != null)
            ws.SheetView.ZoomScaleNormal = (int)UInt32Value.ToUInt32(sheetView.ZoomScaleNormal);
        if (sheetView.ZoomScalePageLayoutView != null)
            ws.SheetView.ZoomScalePageLayoutView = (int)UInt32Value.ToUInt32(sheetView.ZoomScalePageLayoutView);
        if (sheetView.ZoomScaleSheetLayoutView != null)
            ws.SheetView.ZoomScaleSheetLayoutView = (int)UInt32Value.ToUInt32(sheetView.ZoomScaleSheetLayoutView);

        var pane = sheetView.Elements<Pane>().FirstOrDefault();
        if (new[] { PaneStateValues.Frozen, PaneStateValues.FrozenSplit }.Contains(pane?.State?.Value ?? PaneStateValues.Split))
        {
            if (pane.HorizontalSplit != null)
                ws.SheetView.SplitColumn = (Int32)pane.HorizontalSplit.Value;
            if (pane.VerticalSplit != null)
                ws.SheetView.SplitRow = (Int32)pane.VerticalSplit.Value;
        }

        if (XLHelper.IsValidA1Address(sheetView.TopLeftCell))
            ws.SheetView.TopLeftCellAddress = ws.Cell(sheetView.TopLeftCell.Value).Address;
    }

    private static void LoadSheetProtection(SheetProtection sp, XLWorksheet ws)
    {
        if (sp == null) return;

        ws.Protection.IsProtected = OpenXmlHelper.GetBooleanValueAsBool(sp.Sheet, false);

        var algorithmName = sp.AlgorithmName?.Value ?? string.Empty;
        if (String.IsNullOrEmpty(algorithmName))
        {
            ws.Protection.PasswordHash = sp.Password?.Value ?? string.Empty;
            ws.Protection.Base64EncodedSalt = string.Empty;
        }
        else if (DescribedEnumParser<XLProtectionAlgorithm.Algorithm>.IsValidDescription(algorithmName))
        {
            ws.Protection.Algorithm = DescribedEnumParser<XLProtectionAlgorithm.Algorithm>.FromDescription(algorithmName);
            ws.Protection.PasswordHash = sp.HashValue?.Value ?? string.Empty;
            ws.Protection.SpinCount = sp.SpinCount?.Value ?? 0;
            ws.Protection.Base64EncodedSalt = sp.SaltValue?.Value ?? string.Empty;
        }

        ws.Protection.AllowElement(XLSheetProtectionElements.FormatCells, !OpenXmlHelper.GetBooleanValueAsBool(sp.FormatCells, true));
        ws.Protection.AllowElement(XLSheetProtectionElements.FormatColumns, !OpenXmlHelper.GetBooleanValueAsBool(sp.FormatColumns, true));
        ws.Protection.AllowElement(XLSheetProtectionElements.FormatRows, !OpenXmlHelper.GetBooleanValueAsBool(sp.FormatRows, true));
        ws.Protection.AllowElement(XLSheetProtectionElements.InsertColumns, !OpenXmlHelper.GetBooleanValueAsBool(sp.InsertColumns, true));
        ws.Protection.AllowElement(XLSheetProtectionElements.InsertHyperlinks, !OpenXmlHelper.GetBooleanValueAsBool(sp.InsertHyperlinks, true));
        ws.Protection.AllowElement(XLSheetProtectionElements.InsertRows, !OpenXmlHelper.GetBooleanValueAsBool(sp.InsertRows, true));
        ws.Protection.AllowElement(XLSheetProtectionElements.DeleteColumns, !OpenXmlHelper.GetBooleanValueAsBool(sp.DeleteColumns, true));
        ws.Protection.AllowElement(XLSheetProtectionElements.DeleteRows, !OpenXmlHelper.GetBooleanValueAsBool(sp.DeleteRows, true));
        ws.Protection.AllowElement(XLSheetProtectionElements.AutoFilter, !OpenXmlHelper.GetBooleanValueAsBool(sp.AutoFilter, true));
        ws.Protection.AllowElement(XLSheetProtectionElements.PivotTables, !OpenXmlHelper.GetBooleanValueAsBool(sp.PivotTables, true));
        ws.Protection.AllowElement(XLSheetProtectionElements.Sort, !OpenXmlHelper.GetBooleanValueAsBool(sp.Sort, true));
        ws.Protection.AllowElement(XLSheetProtectionElements.EditScenarios, !OpenXmlHelper.GetBooleanValueAsBool(sp.Scenarios, true));

        ws.Protection.AllowElement(XLSheetProtectionElements.EditObjects, !OpenXmlHelper.GetBooleanValueAsBool(sp.Objects, false));
        ws.Protection.AllowElement(XLSheetProtectionElements.SelectLockedCells, !OpenXmlHelper.GetBooleanValueAsBool(sp.SelectLockedCells, false));
        ws.Protection.AllowElement(XLSheetProtectionElements.SelectUnlockedCells, !OpenXmlHelper.GetBooleanValueAsBool(sp.SelectUnlockedCells, false));
    }

    /// <summary>
    /// Loads the conditional formatting.
    /// </summary>
    // https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.conditionalformattingrule%28v=office.15%29.aspx?f=255&MSPPError=-2147217396
    private static void LoadConditionalFormatting(ConditionalFormatting conditionalFormatting, XLWorksheet ws,
        Dictionary<Int32, DifferentialFormat> differentialFormats, LoadContext context)
    {
        if (conditionalFormatting == null) return;

        foreach (var fr in conditionalFormatting.Elements<ConditionalFormattingRule>())
        {
            var ranges = conditionalFormatting.SequenceOfReferences.Items
                .Select(sor => ws.Range(sor.Value));
            var conditionalFormat = new XLConditionalFormat(ranges);

            conditionalFormat.StopIfTrue = OpenXmlHelper.GetBooleanValueAsBool(fr.StopIfTrue, false);

            if (fr.FormatId != null)
            {
                OpenXmlHelper.LoadFont(differentialFormats[(Int32)fr.FormatId.Value].Font, conditionalFormat.Style.Font);
                OpenXmlHelper.LoadFill(differentialFormats[(Int32)fr.FormatId.Value].Fill, conditionalFormat.Style.Fill,
                    differentialFillFormat: true);
                OpenXmlHelper.LoadBorder(differentialFormats[(Int32)fr.FormatId.Value].Border, conditionalFormat.Style.Border);
                OpenXmlHelper.LoadNumberFormat(differentialFormats[(Int32)fr.FormatId.Value].NumberingFormat,
                    conditionalFormat.Style.NumberFormat);
            }

            // The conditional formatting type is compulsory. If it doesn't exist, skip the entire rule.
            if (fr.Type == null) continue;
            conditionalFormat.ConditionalFormatType = fr.Type.Value.ToClosedXml();
            conditionalFormat.Priority = fr.Priority?.Value ?? Int32.MaxValue;

            // Although formulas are directly used only by CellIs and Expression type, other
            // format types also write them for evaluation to the workbook, e.g. rule to
            // IsBlank writes `LEN(TRIM(A2))=0` or ContainsText writes `NOT(ISERROR(SEARCH("hello",A2)))`.
            if (conditionalFormat.ConditionalFormatType == XLConditionalFormatType.CellIs)
            {
                conditionalFormat.Operator = fr.Operator.Value.ToClosedXml();

                // The XML schema allows up to three <formula> tags, but at most two are used.
                // Some producers emit empty <formula> tags that should be ignored and extra
                // non-empty formulas should also be ignored (Excel behavior).
                var nonEmptyFormulas = fr.Elements<Formula>()
                    .Where(static f => !String.IsNullOrEmpty(f.Text))
                    .Select<Formula, XLFormula>(f => GetFormula(f.Text))
                    .ToList();
                if (conditionalFormat.Operator is XLCFOperator.Between or XLCFOperator.NotBetween)
                {
                    var formulas = nonEmptyFormulas.Take(2).ToList();
                    if (formulas.Count != 2)
                        throw PartStructureException.IncorrectElementsCount();

                    conditionalFormat.Values.Add(formulas[0]);
                    conditionalFormat.Values.Add(formulas[1]);
                }
                else
                {
                    // Other XLCFOperators expect one argument.
                    var operatorArg = nonEmptyFormulas.FirstOrDefault();
                    if (operatorArg is null)
                        throw PartStructureException.IncorrectElementsCount();

                    conditionalFormat.Values.Add(operatorArg);
                }
            }
            else if (conditionalFormat.ConditionalFormatType == XLConditionalFormatType.Expression)
            {
                var formula = fr.Elements<Formula>()
                    .Where(static f => !String.IsNullOrEmpty(f.Text))
                    .FirstOrDefault();

                if (formula is null)
                    throw PartStructureException.IncorrectElementsCount();

                conditionalFormat.Values.Add(GetFormula(formula.Text));
            }

            if (!String.IsNullOrWhiteSpace(fr.Text))
                conditionalFormat.Values.Add(GetFormula(fr.Text.Value));

            if (conditionalFormat.ConditionalFormatType == XLConditionalFormatType.Top10)
            {
                if (fr.Percent != null)
                    conditionalFormat.Percent = fr.Percent.Value;
                if (fr.Bottom != null)
                    conditionalFormat.Bottom = fr.Bottom.Value;
                if (fr.Rank != null)
                    conditionalFormat.Values.Add(GetFormula(fr.Rank.Value.ToString()));
            }
            else if (conditionalFormat.ConditionalFormatType == XLConditionalFormatType.TimePeriod)
            {
                if (fr.TimePeriod != null)
                    conditionalFormat.TimePeriod = fr.TimePeriod.Value.ToClosedXml();
                else
                    conditionalFormat.TimePeriod = XLTimePeriod.Yesterday;
            }

            if (fr.Elements<ColorScale>().Any())
            {
                var colorScale = fr.Elements<ColorScale>().First();
                ExtractConditionalFormatValueObjects(conditionalFormat, colorScale);
            }
            else if (fr.Elements<DataBar>().Any())
            {
                var dataBar = fr.Elements<DataBar>().First();
                if (dataBar.ShowValue != null)
                    conditionalFormat.ShowBarOnly = !dataBar.ShowValue.Value;

                var id = fr.Descendants<X14.Id>().FirstOrDefault();
                if (id != null && id.Text != null && !String.IsNullOrWhiteSpace(id.Text))
                    conditionalFormat.Id = new Guid(id.Text.Substring(1, id.Text.Length - 2));

                ExtractConditionalFormatValueObjects(conditionalFormat, dataBar);
            }
            else if (fr.Elements<IconSet>().Any())
            {
                var iconSet = fr.Elements<IconSet>().First();
                if (iconSet.ShowValue != null)
                    conditionalFormat.ShowIconOnly = !iconSet.ShowValue.Value;
                if (iconSet.Reverse != null)
                    conditionalFormat.ReverseIconOrder = iconSet.Reverse.Value;

                if (iconSet.IconSetValue != null)
                    conditionalFormat.IconSetStyle = iconSet.IconSetValue.Value.ToClosedXml();
                else
                    conditionalFormat.IconSetStyle = XLIconSetStyle.ThreeTrafficLights1;

                ExtractConditionalFormatValueObjects(conditionalFormat, iconSet);
            }

            var isPivotTableFormatting = conditionalFormatting.Pivot?.Value ?? false;
            if (isPivotTableFormatting)
                context.AddPivotTableCf(ws.Name, conditionalFormat);
            else
                ws.ConditionalFormats.Add(conditionalFormat);
        }
    }

    private static XLFormula GetFormula(String value)
    {
        var formula = new XLFormula();
        formula._value = value;
        formula.IsFormula = !(value[0] == '"' && value.EndsWith("\""));
        return formula;
    }

    private static void ExtractConditionalFormatValueObjects(XLConditionalFormat conditionalFormat, OpenXmlElement element)
    {
        foreach (var c in element.Elements<ConditionalFormatValueObject>())
        {
            if (c.Type != null)
                conditionalFormat.ContentTypes.Add(c.Type.Value.ToClosedXml());
            if (c.Val != null)
                conditionalFormat.Values.Add(new XLFormula { Value = c.Val.Value });
            else
                conditionalFormat.Values.Add(null);

            if (c.GreaterThanOrEqual != null)
                conditionalFormat.IconSetOperators.Add(c.GreaterThanOrEqual.Value ? XLCFIconSetOperator.EqualOrGreaterThan : XLCFIconSetOperator.GreaterThan);
            else
                conditionalFormat.IconSetOperators.Add(XLCFIconSetOperator.EqualOrGreaterThan);
        }
        foreach (var c in element.Elements<Color>())
        {
            conditionalFormat.Colors.Add(c.ToClosedXMLColor());
        }
    }

    private static void LoadDataValidations(DataValidations dataValidations, XLWorksheet ws)
    {
        if (dataValidations == null) return;

        foreach (DataValidation dvs in dataValidations.Elements<DataValidation>())
        {
            String txt = dvs.SequenceOfReferences.InnerText;
            if (String.IsNullOrWhiteSpace(txt)) continue;
            foreach (var rangeAddress in txt.Split(' '))
            {
                var dvt = new XLDataValidation(ws.Range(rangeAddress));
                ws.DataValidations.Add(dvt, skipIntersectionsCheck: true);
                if (dvs.AllowBlank != null) dvt.IgnoreBlanks = dvs.AllowBlank;
                if (dvs.ShowDropDown != null) dvt.InCellDropdown = !dvs.ShowDropDown.Value;
                if (dvs.ShowErrorMessage != null) dvt.ShowErrorMessage = dvs.ShowErrorMessage;
                if (dvs.ShowInputMessage != null) dvt.ShowInputMessage = dvs.ShowInputMessage;
                if (dvs.PromptTitle != null) dvt.InputTitle = dvs.PromptTitle;
                if (dvs.Prompt != null) dvt.InputMessage = dvs.Prompt;
                if (dvs.ErrorTitle != null) dvt.ErrorTitle = dvs.ErrorTitle;
                if (dvs.Error != null) dvt.ErrorMessage = dvs.Error;
                if (dvs.ErrorStyle != null) dvt.ErrorStyle = dvs.ErrorStyle.Value.ToClosedXml();
                if (dvs.Type != null) dvt.AllowedValues = dvs.Type.Value.ToClosedXml();
                if (dvs.Operator != null) dvt.Operator = dvs.Operator.Value.ToClosedXml();
                if (dvs.Formula1 != null) dvt.MinValue = dvs.Formula1.Text;
                if (dvs.Formula2 != null) dvt.MaxValue = dvs.Formula2.Text;
            }
        }
    }

    private static void LoadHyperlinks(Hyperlinks hyperlinks, WorksheetPart worksheetPart, XLWorksheet ws)
    {
        var hyperlinkDictionary = new Dictionary<String, Uri>();
        if (worksheetPart.HyperlinkRelationships != null)
            hyperlinkDictionary = worksheetPart.HyperlinkRelationships.ToDictionary(hr => hr.Id, hr => hr.Uri);

        if (hyperlinks == null) return;

        foreach (Hyperlink hl in hyperlinks.Elements<Hyperlink>())
        {
            if (hl.Reference.Value.Equals("#REF")) continue;
            String tooltip = hl.Tooltip != null ? hl.Tooltip.Value : String.Empty;
            var xlRange = ws.Range(hl.Reference.Value);
            foreach (XLCell xlCell in xlRange.Cells())
            {
                if (hl.Id != null)
                    xlCell.SetCellHyperlink(new XLHyperlink(hyperlinkDictionary[hl.Id], tooltip));
                else if (hl.Location != null)
                    xlCell.SetCellHyperlink(new XLHyperlink(hl.Location.Value, tooltip));
                else
                    xlCell.SetCellHyperlink(new XLHyperlink(hl.Reference.Value, tooltip));
            }
        }
    }

    private static void LoadPrintOptions(PrintOptions printOptions, XLWorksheet ws)
    {
        if (printOptions == null) return;

        if (printOptions.GridLines != null)
            ws.PageSetup.ShowGridlines = printOptions.GridLines;
        if (printOptions.HorizontalCentered != null)
            ws.PageSetup.CenterHorizontally = printOptions.HorizontalCentered;
        if (printOptions.VerticalCentered != null)
            ws.PageSetup.CenterVertically = printOptions.VerticalCentered;
        if (printOptions.Headings != null)
            ws.PageSetup.ShowRowAndColumnHeadings = printOptions.Headings;
    }

    private static void LoadPageMargins(PageMargins pageMargins, XLWorksheet ws)
    {
        if (pageMargins == null) return;

        if (pageMargins.Bottom != null)
            ws.PageSetup.Margins.Bottom = pageMargins.Bottom;
        if (pageMargins.Footer != null)
            ws.PageSetup.Margins.Footer = pageMargins.Footer;
        if (pageMargins.Header != null)
            ws.PageSetup.Margins.Header = pageMargins.Header;
        if (pageMargins.Left != null)
            ws.PageSetup.Margins.Left = pageMargins.Left;
        if (pageMargins.Right != null)
            ws.PageSetup.Margins.Right = pageMargins.Right;
        if (pageMargins.Top != null)
            ws.PageSetup.Margins.Top = pageMargins.Top;
    }

    private static void LoadPageSetup(PageSetup pageSetup, XLWorksheet ws, PageSetupProperties pageSetupProperties)
    {
        if (pageSetup == null) return;

        if (pageSetup.PaperSize != null)
            ws.PageSetup.PaperSize = (XLPaperSize)Int32.Parse(pageSetup.PaperSize.InnerText);
        if (pageSetup.Scale != null)
            ws.PageSetup.Scale = Int32.Parse(pageSetup.Scale.InnerText);
        if (pageSetupProperties != null && pageSetupProperties.FitToPage != null && pageSetupProperties.FitToPage.Value)
        {
            if (pageSetup.FitToWidth == null)
                ws.PageSetup.PagesWide = 1;
            else
                ws.PageSetup.PagesWide = Int32.Parse(pageSetup.FitToWidth.InnerText);

            if (pageSetup.FitToHeight == null)
                ws.PageSetup.PagesTall = 1;
            else
                ws.PageSetup.PagesTall = Int32.Parse(pageSetup.FitToHeight.InnerText);
        }
        if (pageSetup.PageOrder != null)
            ws.PageSetup.PageOrder = pageSetup.PageOrder.Value.ToClosedXml();
        if (pageSetup.Orientation != null)
            ws.PageSetup.PageOrientation = pageSetup.Orientation.Value.ToClosedXml();
        if (pageSetup.BlackAndWhite != null)
            ws.PageSetup.BlackAndWhite = pageSetup.BlackAndWhite;
        if (pageSetup.Draft != null)
            ws.PageSetup.DraftQuality = pageSetup.Draft;
        if (pageSetup.CellComments != null)
            ws.PageSetup.ShowComments = pageSetup.CellComments.Value.ToClosedXml();
        if (pageSetup.Errors != null)
            ws.PageSetup.PrintErrorValue = pageSetup.Errors.Value.ToClosedXml();
        if (pageSetup.HorizontalDpi != null) ws.PageSetup.HorizontalDpi = (Int32)pageSetup.HorizontalDpi.Value;
        if (pageSetup.VerticalDpi != null) ws.PageSetup.VerticalDpi = (Int32)pageSetup.VerticalDpi.Value;
        if (pageSetup.FirstPageNumber?.HasValue ?? false)
            ws.PageSetup.FirstPageNumber = (int)pageSetup.FirstPageNumber.Value;
    }

    private static void LoadHeaderFooter(HeaderFooter headerFooter, XLWorksheet ws)
    {
        if (headerFooter == null) return;

        if (headerFooter.AlignWithMargins != null)
            ws.PageSetup.AlignHFWithMargins = headerFooter.AlignWithMargins;
        if (headerFooter.ScaleWithDoc != null)
            ws.PageSetup.ScaleHFWithDocument = headerFooter.ScaleWithDoc;

        if (headerFooter.DifferentFirst != null)
            ws.PageSetup.DifferentFirstPageOnHF = headerFooter.DifferentFirst;
        if (headerFooter.DifferentOddEven != null)
            ws.PageSetup.DifferentOddEvenPagesOnHF = headerFooter.DifferentOddEven;

        // Footers
        var xlFooter = (XLHeaderFooter)ws.PageSetup.Footer;
        var evenFooter = headerFooter.EvenFooter;
        if (evenFooter != null)
            xlFooter.SetInnerText(XLHFOccurrence.EvenPages, evenFooter.Text);
        var oddFooter = headerFooter.OddFooter;
        if (oddFooter != null)
            xlFooter.SetInnerText(XLHFOccurrence.OddPages, oddFooter.Text);
        var firstFooter = headerFooter.FirstFooter;
        if (firstFooter != null)
            xlFooter.SetInnerText(XLHFOccurrence.FirstPage, firstFooter.Text);
        // Headers
        var xlHeader = (XLHeaderFooter)ws.PageSetup.Header;
        var evenHeader = headerFooter.EvenHeader;
        if (evenHeader != null)
            xlHeader.SetInnerText(XLHFOccurrence.EvenPages, evenHeader.Text);
        var oddHeader = headerFooter.OddHeader;
        if (oddHeader != null)
            xlHeader.SetInnerText(XLHFOccurrence.OddPages, oddHeader.Text);
        var firstHeader = headerFooter.FirstHeader;
        if (firstHeader != null)
            xlHeader.SetInnerText(XLHFOccurrence.FirstPage, firstHeader.Text);

        ((XLHeaderFooter)ws.PageSetup.Header).SetAsInitial();
        ((XLHeaderFooter)ws.PageSetup.Footer).SetAsInitial();
    }

    private static void LoadRowBreaks(RowBreaks rowBreaks, XLWorksheet ws)
    {
        if (rowBreaks == null) return;

        foreach (Break rowBreak in rowBreaks.Elements<Break>())
            ws.PageSetup.RowBreaks.Add(Int32.Parse(rowBreak.Id.InnerText));
    }

    private static void LoadColumnBreaks(ColumnBreaks columnBreaks, XLWorksheet ws)
    {
        if (columnBreaks == null) return;

        foreach (Break columnBreak in columnBreaks.Elements<Break>().Where(columnBreak => columnBreak.Id != null))
        {
            ws.PageSetup.ColumnBreaks.Add(Int32.Parse(columnBreak.Id.InnerText));
        }
    }

    private static void LoadExtensions(WorksheetExtensionList extensions, XLWorksheet ws)
    {
        if (extensions == null)
        {
            return;
        }

        foreach (var dvs in extensions
                     .Descendants<X14.DataValidations>()
                     .SelectMany(dataValidations => dataValidations.Descendants<X14.DataValidation>()))
        {
            String txt = dvs.ReferenceSequence.InnerText;
            if (String.IsNullOrWhiteSpace(txt)) continue;
            foreach (var rangeAddress in txt.Split(' '))
            {
                var dvt = new XLDataValidation(ws.Range(rangeAddress));
                ws.DataValidations.Add(dvt, skipIntersectionsCheck: true);
                if (dvs.AllowBlank != null) dvt.IgnoreBlanks = dvs.AllowBlank;
                if (dvs.ShowDropDown != null) dvt.InCellDropdown = !dvs.ShowDropDown.Value;
                if (dvs.ShowErrorMessage != null) dvt.ShowErrorMessage = dvs.ShowErrorMessage;
                if (dvs.ShowInputMessage != null) dvt.ShowInputMessage = dvs.ShowInputMessage;
                if (dvs.PromptTitle != null) dvt.InputTitle = dvs.PromptTitle;
                if (dvs.Prompt != null) dvt.InputMessage = dvs.Prompt;
                if (dvs.ErrorTitle != null) dvt.ErrorTitle = dvs.ErrorTitle;
                if (dvs.Error != null) dvt.ErrorMessage = dvs.Error;
                if (dvs.ErrorStyle != null) dvt.ErrorStyle = dvs.ErrorStyle.Value.ToClosedXml();
                if (dvs.Type != null) dvt.AllowedValues = dvs.Type.Value.ToClosedXml();
                if (dvs.Operator != null) dvt.Operator = dvs.Operator.Value.ToClosedXml();
                if (dvs.DataValidationForumla1 != null) dvt.MinValue = dvs.DataValidationForumla1.InnerText;
                if (dvs.DataValidationForumla2 != null) dvt.MaxValue = dvs.DataValidationForumla2.InnerText;
            }
        }

        foreach (var conditionalFormattingRule in extensions
                     .Descendants<X14.ConditionalFormattingRule>()
                     .Where(cf =>
                         cf.Type != null
                         && cf.Type.HasValue
                         && cf.Type.Value == ConditionalFormatValues.DataBar))
        {
            var xlConditionalFormat = ws.ConditionalFormats
                .Cast<XLConditionalFormat>()
                .SingleOrDefault(cf => cf.Id.WrapInBraces() == conditionalFormattingRule.Id);
            if (xlConditionalFormat != null)
            {
                var negativeFillColor = conditionalFormattingRule.Descendants<X14.NegativeFillColor>().SingleOrDefault();
                xlConditionalFormat.Colors.Add(negativeFillColor.ToClosedXMLColor());
            }
        }

        foreach (var slg in extensions
                     .Descendants<X14.SparklineGroups>()
                     .SelectMany(sparklineGroups => sparklineGroups.Descendants<X14.SparklineGroup>()))
        {
            var xlSparklineGroup = (ws.SparklineGroups as XLSparklineGroups).Add();

            if (slg.Formula != null)
                xlSparklineGroup.DateRange = ws.Workbook.Range(slg.Formula.Text);

            var xlSparklineStyle = xlSparklineGroup.Style;
            if (slg.FirstMarkerColor != null) xlSparklineStyle.FirstMarkerColor = slg.FirstMarkerColor.ToClosedXMLColor();
            if (slg.LastMarkerColor != null) xlSparklineStyle.LastMarkerColor = slg.LastMarkerColor.ToClosedXMLColor();
            if (slg.HighMarkerColor != null) xlSparklineStyle.HighMarkerColor = slg.HighMarkerColor.ToClosedXMLColor();
            if (slg.LowMarkerColor != null) xlSparklineStyle.LowMarkerColor = slg.LowMarkerColor.ToClosedXMLColor();
            if (slg.SeriesColor != null) xlSparklineStyle.SeriesColor = slg.SeriesColor.ToClosedXMLColor();
            if (slg.NegativeColor != null) xlSparklineStyle.NegativeColor = slg.NegativeColor.ToClosedXMLColor();
            if (slg.MarkersColor != null) xlSparklineStyle.MarkersColor = slg.MarkersColor.ToClosedXMLColor();
            xlSparklineGroup.Style = xlSparklineStyle;

            if (slg.DisplayHidden != null) xlSparklineGroup.DisplayHidden = slg.DisplayHidden;
            if (slg.LineWeight != null) xlSparklineGroup.LineWeight = slg.LineWeight;
            if (slg.Type != null) xlSparklineGroup.Type = slg.Type.Value.ToClosedXml();
            if (slg.DisplayEmptyCellsAs != null) xlSparklineGroup.DisplayEmptyCellsAs = slg.DisplayEmptyCellsAs.Value.ToClosedXml();

            xlSparklineGroup.ShowMarkers = XLSparklineMarkers.None;
            if (OpenXmlHelper.GetBooleanValueAsBool(slg.Markers, false)) xlSparklineGroup.ShowMarkers |= XLSparklineMarkers.Markers;
            if (OpenXmlHelper.GetBooleanValueAsBool(slg.High, false)) xlSparklineGroup.ShowMarkers |= XLSparklineMarkers.HighPoint;
            if (OpenXmlHelper.GetBooleanValueAsBool(slg.Low, false)) xlSparklineGroup.ShowMarkers |= XLSparklineMarkers.LowPoint;
            if (OpenXmlHelper.GetBooleanValueAsBool(slg.First, false)) xlSparklineGroup.ShowMarkers |= XLSparklineMarkers.FirstPoint;
            if (OpenXmlHelper.GetBooleanValueAsBool(slg.Last, false)) xlSparklineGroup.ShowMarkers |= XLSparklineMarkers.LastPoint;
            if (OpenXmlHelper.GetBooleanValueAsBool(slg.Negative, false)) xlSparklineGroup.ShowMarkers |= XLSparklineMarkers.NegativePoints;

            if (slg.AxisColor != null) xlSparklineGroup.HorizontalAxis.Color = XLColor.FromHtml(slg.AxisColor.Rgb.Value);
            if (slg.DisplayXAxis != null) xlSparklineGroup.HorizontalAxis.IsVisible = slg.DisplayXAxis;
            if (slg.RightToLeft != null) xlSparklineGroup.HorizontalAxis.RightToLeft = slg.RightToLeft;

            if (slg.ManualMax != null) xlSparklineGroup.VerticalAxis.ManualMax = slg.ManualMax;
            if (slg.ManualMin != null) xlSparklineGroup.VerticalAxis.ManualMin = slg.ManualMin;
            if (slg.MinAxisType != null) xlSparklineGroup.VerticalAxis.MinAxisType = slg.MinAxisType.Value.ToClosedXml();
            if (slg.MaxAxisType != null) xlSparklineGroup.VerticalAxis.MaxAxisType = slg.MaxAxisType.Value.ToClosedXml();

            slg.Descendants<X14.Sparklines>().SelectMany(sls => sls.Descendants<X14.Sparkline>())
                .ForEach(sl => xlSparklineGroup.Add(sl.ReferenceSequence?.Text, sl.Formula?.Text));
        }
    }

    private static void ApplyStyle(IXLStylized xlStylized, Int32 styleIndex, Stylesheet s, Fills fills, Borders borders,
        NumberingFormats numberingFormats, XLWorkbookStyles styles)
    {
        var xlStyleKey = XLStyle.Default.Key;
        XLWorkbook.LoadStyle(ref xlStyleKey, styleIndex, s, fills, borders, numberingFormats, styles);

        // When loading columns we must propagate style to each column but not deeper. In other cases we do not propagate at all.
        if (xlStylized is IXLColumns columns)
        {
            columns.Cast<XLColumn>().ForEach(col => col.InnerStyle = new XLStyle(col, xlStyleKey));
        }
        else
        {
            xlStylized.InnerStyle = new XLStyle(xlStylized, xlStyleKey);
        }
    }
}
