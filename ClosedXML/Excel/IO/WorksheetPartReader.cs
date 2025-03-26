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

namespace ClosedXML.Excel.IO;

#nullable disable

internal class WorksheetPartReader
{
    public static void LoadSheetProperties(SheetProperties sheetProperty, XLWorksheet ws, out PageSetupProperties pageSetupProperties)
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

    internal static void LoadColumns(Stylesheet s, NumberingFormats numberingFormats, Fills fills, Borders borders,
                             Fonts fonts, XLWorksheet ws, Columns columns)
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
            XLWorkbook.ApplyStyle(ws, styleIndexDefault, s, fills, borders, fonts, numberingFormats);

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
                XLWorkbook.ApplyStyle(xlColumns, styleIndex, s, fills, borders, fonts, numberingFormats);
            }
            else
            {
                xlColumns.Style = ws.Style;
            }
        }
    }

    public static void LoadSheetViews(SheetViews sheetViews, XLWorksheet ws)
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

    public static void LoadSheetProtection(SheetProtection sp, XLWorksheet ws)
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

    public static void LoadAutoFilter(AutoFilter af, XLWorksheet ws)
    {
        if (af != null)
        {
            ws.Range(af.Reference.Value).SetAutoFilter();
            var autoFilter = ws.AutoFilter;
            LoadAutoFilterSort(af, ws, autoFilter);
            LoadAutoFilterColumns(af, autoFilter);
        }
    }

    public static void LoadAutoFilterColumns(AutoFilter af, XLAutoFilter autoFilter)
    {
        foreach (var filterColumn in af.Elements<FilterColumn>())
        {
            Int32 column = (int)filterColumn.ColumnId.Value + 1;
            var xlFilterColumn = autoFilter.Column(column);
            if (filterColumn.CustomFilters is { } customFilters)
            {
                xlFilterColumn.FilterType = XLFilterType.Custom;
                var connector = OpenXmlHelper.GetBooleanValueAsBool(customFilters.And, false) ? XLConnector.And : XLConnector.Or;

                foreach (var filter in customFilters.OfType<CustomFilter>())
                {
                    // Equal or NotEqual use wildcards, not value comparison. The rest does value comparison.
                    // There is no filter operation for equal of numbers (maybe combine >= and <=).
                    var op = filter.Operator is not null ? filter.Operator.Value.ToClosedXml() : XLFilterOperator.Equal;
                    XLFilter xlFilter;
                    var filterValue = filter.Val.Value;
                    switch (op)
                    {
                        case XLFilterOperator.Equal:
                            xlFilter = XLFilter.CreateCustomPatternFilter(filterValue, true, connector);
                            break;
                        case XLFilterOperator.NotEqual:
                            xlFilter = XLFilter.CreateCustomPatternFilter(filterValue, false, connector);
                            break;
                        default:
                            // OOXML allows only string, so do your best to convert back to a properly typed
                            // variable. It's not perfect, but let's mimic Excel.
                            var customValue = XLCellValue.FromText(filterValue, CultureInfo.InvariantCulture);
                            xlFilter = XLFilter.CreateCustomFilter(customValue, op, connector);
                            break;
                    }

                    xlFilterColumn.AddFilter(xlFilter);
                }
            }
            else if (filterColumn.Filters is { } filters)
            {
                xlFilterColumn.FilterType = XLFilterType.Regular;
                foreach (var filter in filters.OfType<Filter>())
                {
                    xlFilterColumn.AddFilter(XLFilter.CreateRegularFilter(filter.Val.Value));
                }

                foreach (var dateGroupItem in filters.OfType<DateGroupItem>())
                {
                    if (dateGroupItem.DateTimeGrouping is null || !dateGroupItem.DateTimeGrouping.HasValue)
                        continue;

                    var xlGrouping = dateGroupItem.DateTimeGrouping.Value.ToClosedXml();
                    var year = 1900;
                    var month = 1;
                    var day = 1;
                    var hour = 0;
                    var minute = 0;
                    var second = 0;

                    var valid = true;

                    if (xlGrouping >= XLDateTimeGrouping.Year)
                    {
                        if (dateGroupItem.Year?.HasValue ?? false)
                            year = dateGroupItem.Year.Value;
                        else
                            valid = false;
                    }

                    if (xlGrouping >= XLDateTimeGrouping.Month)
                    {
                        if (dateGroupItem.Month?.HasValue ?? false)
                            month = dateGroupItem.Month.Value;
                        else
                            valid = false;
                    }

                    if (xlGrouping >= XLDateTimeGrouping.Day)
                    {
                        if (dateGroupItem.Day?.HasValue ?? false)
                            day = dateGroupItem.Day.Value;
                        else
                            valid = false;
                    }

                    if (xlGrouping >= XLDateTimeGrouping.Hour)
                    {
                        if (dateGroupItem.Hour?.HasValue ?? false)
                            hour = dateGroupItem.Hour.Value;
                        else
                            valid = false;
                    }

                    if (xlGrouping >= XLDateTimeGrouping.Minute)
                    {
                        if (dateGroupItem.Minute?.HasValue ?? false)
                            minute = dateGroupItem.Minute.Value;
                        else
                            valid = false;
                    }

                    if (xlGrouping >= XLDateTimeGrouping.Second)
                    {
                        if (dateGroupItem.Second?.HasValue ?? false)
                            second = dateGroupItem.Second.Value;
                        else
                            valid = false;
                    }

                    if (valid)
                    {
                        var date = new DateTime(year, month, day, hour, minute, second);
                        var xlDateGroupFilter = XLFilter.CreateDateGroupFilter(date, xlGrouping);
                        xlFilterColumn.AddFilter(xlDateGroupFilter);
                    }
                }
            }
            else if (filterColumn.Top10 is { } top10)
            {
                xlFilterColumn.FilterType = XLFilterType.TopBottom;
                xlFilterColumn.TopBottomType = OpenXmlHelper.GetBooleanValueAsBool(top10.Percent, false)
                    ? XLTopBottomType.Percent
                    : XLTopBottomType.Items;
                var takeTop = OpenXmlHelper.GetBooleanValueAsBool(top10.Top, true);
                xlFilterColumn.TopBottomPart = takeTop ? XLTopBottomPart.Top : XLTopBottomPart.Bottom;

                // Value contains how many percent or items, so it can only be int.
                // Filter value is optional, so we don't rely on it.
                var percentsOrItems = (int)top10.Val.Value;
                xlFilterColumn.TopBottomValue = percentsOrItems;
                xlFilterColumn.AddFilter(XLFilter.CreateTopBottom(takeTop, percentsOrItems));
            }
            else if (filterColumn.DynamicFilter is { } dynamicFilter)
            {
                xlFilterColumn.FilterType = XLFilterType.Dynamic;
                var dynamicType = dynamicFilter.Type is { } dynamicFilterType
                    ? dynamicFilterType.Value.ToClosedXml()
                    : XLFilterDynamicType.AboveAverage;
                var dynamicValue = filterColumn.DynamicFilter.Val.Value;

                xlFilterColumn.DynamicType = dynamicType;
                xlFilterColumn.DynamicValue = dynamicValue;
                xlFilterColumn.AddFilter(XLFilter.CreateAverage(dynamicValue, dynamicType == XLFilterDynamicType.AboveAverage));
            }
        }
    }

    private static void LoadAutoFilterSort(AutoFilter af, XLWorksheet ws, XLAutoFilter autoFilter)
    {
        var sort = af.Elements<SortState>().FirstOrDefault();
        if (sort != null)
        {
            var condition = sort.Elements<SortCondition>().FirstOrDefault();
            if (condition != null)
            {
                Int32 column = ws.Range(condition.Reference.Value).FirstCell().Address.ColumnNumber - autoFilter.Range.FirstCell().Address.ColumnNumber + 1;
                autoFilter.SortColumn = column;
                autoFilter.Sorted = true;
                autoFilter.SortOrder = condition.Descending != null && condition.Descending.Value ? XLSortOrder.Descending : XLSortOrder.Ascending;
            }
        }
    }

    /// <summary>
    /// Loads the conditional formatting.
    /// </summary>
    // https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.conditionalformattingrule%28v=office.15%29.aspx?f=255&MSPPError=-2147217396
    public static void LoadConditionalFormatting(ConditionalFormatting conditionalFormatting, XLWorksheet ws,
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
        foreach (var c in element.Elements<DocumentFormat.OpenXml.Spreadsheet.Color>())
        {
            conditionalFormat.Colors.Add(c.ToClosedXMLColor());
        }
    }

    public static void LoadDataValidations(DataValidations dataValidations, XLWorksheet ws)
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

    public static void LoadHyperlinks(Hyperlinks hyperlinks, WorksheetPart worksheetPart, XLWorksheet ws)
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

    public static void LoadPrintOptions(PrintOptions printOptions, XLWorksheet ws)
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

    public static void LoadPageMargins(PageMargins pageMargins, XLWorksheet ws)
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

    public static void LoadPageSetup(PageSetup pageSetup, XLWorksheet ws, PageSetupProperties pageSetupProperties)
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

    public static void LoadHeaderFooter(HeaderFooter headerFooter, XLWorksheet ws)
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

    public static void LoadRowBreaks(RowBreaks rowBreaks, XLWorksheet ws)
    {
        if (rowBreaks == null) return;

        foreach (Break rowBreak in rowBreaks.Elements<Break>())
            ws.PageSetup.RowBreaks.Add(Int32.Parse(rowBreak.Id.InnerText));
    }

    public static void LoadColumnBreaks(ColumnBreaks columnBreaks, XLWorksheet ws)
    {
        if (columnBreaks == null) return;

        foreach (Break columnBreak in columnBreaks.Elements<Break>().Where(columnBreak => columnBreak.Id != null))
        {
            ws.PageSetup.ColumnBreaks.Add(Int32.Parse(columnBreak.Id.InnerText));
        }
    }

    public static void LoadExtensions(WorksheetExtensionList extensions, XLWorksheet ws)
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
}
