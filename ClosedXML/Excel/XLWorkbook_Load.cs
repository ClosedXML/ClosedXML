using ClosedXML.Extensions;
using ClosedXML.Utils;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Op = DocumentFormat.OpenXml.CustomProperties;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace ClosedXML.Excel
{
    using Ap;
    using Drawings;
    using Op;
    using System.Drawing;

    public partial class XLWorkbook
    {
        private readonly Dictionary<String, Color> _colorList = new Dictionary<string, Color>();

        private void Load(String file)
        {
            LoadSheets(file);
        }

        private void Load(Stream stream)
        {
            LoadSheets(stream);
        }

        private void LoadSheets(String fileName)
        {
            using (var dSpreadsheet = SpreadsheetDocument.Open(fileName, false))
                LoadSpreadsheetDocument(dSpreadsheet);
        }

        private void LoadSheets(Stream stream)
        {
            using (var dSpreadsheet = SpreadsheetDocument.Open(stream, false))
                LoadSpreadsheetDocument(dSpreadsheet);
        }

        private void LoadSheetsFromTemplate(String fileName)
        {
            using (var dSpreadsheet = SpreadsheetDocument.CreateFromTemplate(fileName))
                LoadSpreadsheetDocument(dSpreadsheet);
        }

        private void LoadSpreadsheetDocument(SpreadsheetDocument dSpreadsheet)
        {
            ShapeIdManager = new XLIdManager();
            SetProperties(dSpreadsheet);

            SharedStringItem[] sharedStrings = null;
            if (dSpreadsheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Any())
            {
                var shareStringPart = dSpreadsheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                sharedStrings = shareStringPart.SharedStringTable.Elements<SharedStringItem>().ToArray();
            }

            if (dSpreadsheet.CustomFilePropertiesPart != null)
            {
                foreach (var m in dSpreadsheet.CustomFilePropertiesPart.Properties.Elements<CustomDocumentProperty>())
                {
                    String name = m.Name?.Value;

                    if (string.IsNullOrWhiteSpace(name))
                        continue;

                    if (m.VTLPWSTR != null)
                        CustomProperties.Add(name, m.VTLPWSTR.Text);
                    else if (m.VTFileTime != null)
                    {
                        CustomProperties.Add(name,
                                             DateTime.ParseExact(m.VTFileTime.Text, "yyyy'-'MM'-'dd'T'HH':'mm':'ssK",
                                                                 CultureInfo.InvariantCulture));
                    }
                    else if (m.VTDouble != null)
                        CustomProperties.Add(name, Double.Parse(m.VTDouble.Text, CultureInfo.InvariantCulture));
                    else if (m.VTBool != null)
                        CustomProperties.Add(name, m.VTBool.Text == "true");
                }
            }

            var wbProps = dSpreadsheet.WorkbookPart.Workbook.WorkbookProperties;
            Use1904DateSystem = wbProps?.Date1904?.Value ?? false;

            var wbProtection = dSpreadsheet.WorkbookPart.Workbook.WorkbookProtection;
            if (wbProtection != null)
            {
                if (wbProtection.LockStructure != null)
                    LockStructure = wbProtection.LockStructure.Value;
                if (wbProtection.LockWindows != null)
                    LockWindows = wbProtection.LockWindows.Value;
                if (wbProtection.WorkbookPassword != null)
                    LockPassword = wbProtection.WorkbookPassword.Value;
            }

            var calculationProperties = dSpreadsheet.WorkbookPart.Workbook.CalculationProperties;
            if (calculationProperties != null)
            {
                var calculateMode = calculationProperties.CalculationMode;
                if (calculateMode != null)
                    CalculateMode = calculateMode.Value.ToClosedXml();

                var calculationOnSave = calculationProperties.CalculationOnSave;
                if (calculationOnSave != null)
                    CalculationOnSave = calculationOnSave.Value;

                var forceFullCalculation = calculationProperties.ForceFullCalculation;
                if (forceFullCalculation != null)
                    ForceFullCalculation = forceFullCalculation.Value;

                var fullCalculationOnLoad = calculationProperties.FullCalculationOnLoad;
                if (fullCalculationOnLoad != null)
                    FullCalculationOnLoad = fullCalculationOnLoad.Value;

                var fullPrecision = calculationProperties.FullPrecision;
                if (fullPrecision != null)
                    FullPrecision = fullPrecision.Value;

                var referenceMode = calculationProperties.ReferenceMode;
                if (referenceMode != null)
                    ReferenceStyle = referenceMode.Value.ToClosedXml();
            }

            var efp = dSpreadsheet.ExtendedFilePropertiesPart;
            if (efp != null && efp.Properties != null)
            {
                if (efp.Properties.Elements<Company>().Any())
                    Properties.Company = efp.Properties.GetFirstChild<Company>().Text;

                if (efp.Properties.Elements<Manager>().Any())
                    Properties.Manager = efp.Properties.GetFirstChild<Manager>().Text;
            }

            Stylesheet s = null;
            if (dSpreadsheet.WorkbookPart.WorkbookStylesPart != null &&
                dSpreadsheet.WorkbookPart.WorkbookStylesPart.Stylesheet != null)
            {
                s = dSpreadsheet.WorkbookPart.WorkbookStylesPart.Stylesheet;
            }

            NumberingFormats numberingFormats = s == null ? null : s.NumberingFormats;
            Fills fills = s == null ? null : s.Fills;
            Borders borders = s == null ? null : s.Borders;
            Fonts fonts = s == null ? null : s.Fonts;
            Int32 dfCount = 0;
            Dictionary<Int32, DifferentialFormat> differentialFormats;
            if (s != null && s.DifferentialFormats != null)
                differentialFormats = s.DifferentialFormats.Elements<DifferentialFormat>().ToDictionary(k => dfCount++);
            else
                differentialFormats = new Dictionary<Int32, DifferentialFormat>();

            var sheets = dSpreadsheet.WorkbookPart.Workbook.Sheets;
            Int32 position = 0;
            foreach (var dSheet in sheets.OfType<Sheet>())
            {
                position++;
                var sharedFormulasR1C1 = new Dictionary<UInt32, String>();

                var worksheetPart = dSpreadsheet.WorkbookPart.GetPartById(dSheet.Id) as WorksheetPart;

                if (worksheetPart == null)
                {
                    UnsupportedSheets.Add(new UnsupportedSheet { SheetId = dSheet.SheetId.Value, Position = position });
                    continue;
                }

                var sheetName = dSheet.Name;

                var ws = (XLWorksheet)WorksheetsInternal.Add(sheetName, position);
                ws.RelId = dSheet.Id;
                ws.SheetId = (Int32)dSheet.SheetId.Value;

                if (dSheet.State != null)
                    ws.Visibility = dSheet.State.Value.ToClosedXml();

                var styleList = new Dictionary<int, IXLStyle>();// {{0, ws.Style}};
                PageSetupProperties pageSetupProperties = null;

                using (var reader = OpenXmlReader.Create(worksheetPart))
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
                                {
                                    ws.ColumnWidth = sheetFormatProperties.DefaultColumnWidth;
                                }
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
                            LoadColumns(s, numberingFormats, fills, borders, fonts, ws,
                                        (Columns)reader.LoadCurrentElement());
                        else if (reader.ElementType == typeof(Row))
                        {
                            lastRow = 0;
                            LoadRows(s, numberingFormats, fills, borders, fonts, ws, sharedStrings, sharedFormulasR1C1,
                                     styleList, (Row)reader.LoadCurrentElement());
                        }
                        else if (reader.ElementType == typeof(AutoFilter))
                            LoadAutoFilter((AutoFilter)reader.LoadCurrentElement(), ws);
                        else if (reader.ElementType == typeof(SheetProtection))
                            LoadSheetProtection((SheetProtection)reader.LoadCurrentElement(), ws);
                        else if (reader.ElementType == typeof(DataValidations))
                            LoadDataValidations((DataValidations)reader.LoadCurrentElement(), ws);
                        else if (reader.ElementType == typeof(ConditionalFormatting))
                            LoadConditionalFormatting((ConditionalFormatting)reader.LoadCurrentElement(), ws, differentialFormats);
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

                (ws.ConditionalFormats as XLConditionalFormats).ReorderAccordingToOriginalPriority();

                #region LoadTables

                foreach (var tableDefinitionPart in worksheetPart.TableDefinitionParts)
                {
                    var relId = worksheetPart.GetIdOfPart(tableDefinitionPart);
                    var dTable = tableDefinitionPart.Table;

                    String reference = dTable.Reference.Value;
                    String tableName = dTable.Name ?? dTable.DisplayName ?? string.Empty;
                    if (String.IsNullOrWhiteSpace(tableName))
                        throw new InvalidDataException("The table name is missing.");

                    var xlTable = ws.Range(reference).CreateTable(tableName, false) as XLTable;
                    xlTable.RelId = relId;

                    if (dTable.HeaderRowCount != null && dTable.HeaderRowCount == 0)
                    {
                        xlTable._showHeaderRow = false;
                        //foreach (var tableColumn in dTable.TableColumns.Cast<TableColumn>())
                        xlTable.AddFields(dTable.TableColumns.Cast<TableColumn>().Select(t => GetTableColumnName(t.Name.Value)));
                    }
                    else
                    {
                        xlTable.InitializeAutoFilter();
                    }

                    if (dTable.TotalsRowCount != null && dTable.TotalsRowCount.Value > 0)
                        ((XLTable)xlTable)._showTotalsRow = true;

                    if (dTable.TableStyleInfo != null)
                    {
                        if (dTable.TableStyleInfo.ShowFirstColumn != null)
                            xlTable.EmphasizeFirstColumn = dTable.TableStyleInfo.ShowFirstColumn.Value;
                        if (dTable.TableStyleInfo.ShowLastColumn != null)
                            xlTable.EmphasizeLastColumn = dTable.TableStyleInfo.ShowLastColumn.Value;
                        if (dTable.TableStyleInfo.ShowRowStripes != null)
                            xlTable.ShowRowStripes = dTable.TableStyleInfo.ShowRowStripes.Value;
                        if (dTable.TableStyleInfo.ShowColumnStripes != null)
                            xlTable.ShowColumnStripes = dTable.TableStyleInfo.ShowColumnStripes.Value;
                        if (dTable.TableStyleInfo.Name != null)
                        {
                            var theme = XLTableTheme.FromName(dTable.TableStyleInfo.Name.Value);
                            if (theme != null)
                                xlTable.Theme = theme;
                            else
                                xlTable.Theme = new XLTableTheme(dTable.TableStyleInfo.Name.Value);
                        }
                        else
                            xlTable.Theme = XLTableTheme.None;
                    }

                    if (dTable.AutoFilter != null)
                    {
                        xlTable.ShowAutoFilter = true;
                        LoadAutoFilterColumns(dTable.AutoFilter, xlTable.AutoFilter);
                    }
                    else
                        xlTable.ShowAutoFilter = false;

                    if (xlTable.ShowTotalsRow)
                    {
                        foreach (var tableColumn in dTable.TableColumns.Cast<TableColumn>())
                        {
                            var tableColumnName = GetTableColumnName(tableColumn.Name.Value);
                            if (tableColumn.TotalsRowFunction != null)
                                xlTable.Field(tableColumnName).TotalsRowFunction =
                                    tableColumn.TotalsRowFunction.Value.ToClosedXml();

                            if (tableColumn.TotalsRowFormula != null)
                                xlTable.Field(tableColumnName).TotalsRowFormulaA1 =
                                    tableColumn.TotalsRowFormula.Text;

                            if (tableColumn.TotalsRowLabel != null)
                                xlTable.Field(tableColumnName).TotalsRowLabel = tableColumn.TotalsRowLabel.Value;
                        }
                        if (xlTable.AutoFilter != null)
                            xlTable.AutoFilter.Range = xlTable.Worksheet.Range(
                                                    xlTable.RangeAddress.FirstAddress.RowNumber, xlTable.RangeAddress.FirstAddress.ColumnNumber,
                                                    xlTable.RangeAddress.LastAddress.RowNumber - 1, xlTable.RangeAddress.LastAddress.ColumnNumber);
                    }
                    else if (xlTable.AutoFilter != null)
                        xlTable.AutoFilter.Range = xlTable.Worksheet.Range(xlTable.RangeAddress);
                }

                #endregion LoadTables

                LoadDrawings(worksheetPart, ws);

                #region LoadComments

                if (worksheetPart.WorksheetCommentsPart != null)
                {
                    var root = worksheetPart.WorksheetCommentsPart.Comments;
                    var authors = root.GetFirstChild<Authors>().ChildElements;
                    var comments = root.GetFirstChild<CommentList>().ChildElements;

                    // **** MAYBE FUTURE SHAPE SIZE SUPPORT
                    XDocument xdoc = GetCommentVmlFile(worksheetPart);

                    foreach (Comment c in comments)
                    {
                        // find cell by reference
                        var cell = ws.Cell(c.Reference);

                        var xlComment = cell.Comment;
                        xlComment.Author = authors[(int)c.AuthorId.Value].InnerText;
                        //xlComment.ShapeId = (Int32)c.ShapeId.Value;
                        //ShapeIdManager.Add(xlComment.ShapeId);

                        var runs = c.GetFirstChild<CommentText>().Elements<Run>();
                        foreach (Run run in runs)
                        {
                            var runProperties = run.RunProperties;
                            String text = run.Text.InnerText.FixNewLines();
                            var rt = xlComment.AddText(text);
                            LoadFont(runProperties, rt);
                        }

                        XElement shape = GetCommentShape(xdoc);

                        LoadShapeProperties<IXLComment>(xlComment, shape);

                        var clientData = shape.Elements().First(e => e.Name.LocalName == "ClientData");
                        LoadClientData<IXLComment>(xlComment, clientData);

                        var textBox = shape.Elements().First(e => e.Name.LocalName == "textbox");
                        LoadTextBox<IXLComment>(xlComment, textBox);

                        var alt = shape.Attribute("alt");
                        if (alt != null) xlComment.Style.Web.SetAlternateText(alt.Value);

                        LoadColorsAndLines<IXLComment>(xlComment, shape);

                        //var insetmode = (string)shape.Attributes().First(a=> a.Name.LocalName == "insetmode");
                        //xlComment.Style.Margins.Automatic = insetmode != null && insetmode.Equals("auto");

                        shape.Remove();
                    }
                }

                #endregion LoadComments
            }

            var workbook = dSpreadsheet.WorkbookPart.Workbook;

            var bookViews = workbook.BookViews;
            if (bookViews != null && bookViews.Any())
            {
                var workbookView = bookViews.First() as WorkbookView;
                if (workbookView != null && workbookView.ActiveTab != null)
                {
                    UnsupportedSheet unsupportedSheet =
                        UnsupportedSheets.FirstOrDefault(us => us.Position == (Int32)(workbookView.ActiveTab.Value + 1));
                    if (unsupportedSheet != null)
                        unsupportedSheet.IsActive = true;
                    else
                    {
                        Worksheet((Int32)(workbookView.ActiveTab.Value + 1)).SetTabActive();
                    }
                }
            }
            LoadDefinedNames(workbook);

            #region Pivot tables

            // Delay loading of pivot tables until all sheets have been loaded
            foreach (var dSheet in sheets.OfType<Sheet>())
            {
                var worksheetPart = dSpreadsheet.WorkbookPart.GetPartById(dSheet.Id) as WorksheetPart;

                if (worksheetPart != null)
                {
                    var ws = (XLWorksheet)WorksheetsInternal.Worksheet(dSheet.Name);

                    foreach (var pivotTablePart in worksheetPart.PivotTableParts)
                    {
                        var pivotTableCacheDefinitionPart = pivotTablePart.PivotTableCacheDefinitionPart;
                        var pivotTableDefinition = pivotTablePart.PivotTableDefinition;

                        var target = ws.FirstCell();
                        if (pivotTableDefinition.Location != null && pivotTableDefinition.Location.Reference != null && pivotTableDefinition.Location.Reference.HasValue)
                        {
                            target = ws.Range(pivotTableDefinition.Location.Reference.Value).FirstCell();
                        }

                        IXLRange source = null;
                        XLPivotTableSourceType sourceType = XLPivotTableSourceType.Range;
                        if (pivotTableCacheDefinitionPart?.PivotCacheDefinition?.CacheSource?.WorksheetSource != null)
                        {
                            // TODO: Implement other sources besides worksheetSource
                            // But for now assume names and references point directly to a range
                            var wss = pivotTableCacheDefinitionPart.PivotCacheDefinition.CacheSource.WorksheetSource;

                            if (!String.IsNullOrEmpty(wss.Id))
                            {
                                var externalRelationship = pivotTableCacheDefinitionPart.ExternalRelationships.FirstOrDefault(er => er.Id.Equals(wss.Id));
                                if (externalRelationship?.IsExternal ?? false)
                                {
                                    // We don't support external sources
                                    continue;
                                }
                            }

                            if (wss.Name != null)
                            {
                                var table = ws
                                    .Workbook
                                    .Worksheets
                                    .SelectMany(ws1 => ws1.Tables)
                                    .FirstOrDefault(t => t.Name.Equals(wss.Name.Value));

                                if (table != null)
                                {
                                    sourceType = XLPivotTableSourceType.Table;
                                    source = table;
                                }
                                else
                                {
                                    sourceType = XLPivotTableSourceType.Range;
                                    source = this.Range(wss.Name.Value);
                                }
                            }
                            else
                            {
                                sourceType = XLPivotTableSourceType.Range;
                                var sourceSheet = wss.Sheet == null ? ws : this.Worksheet(wss.Sheet.Value);
                                source = this.Range(sourceSheet.Range(wss.Reference.Value).RangeAddress.ToStringRelative(includeSheet: true));
                            }

                            if (source == null)
                                continue;
                        }

                        if (target != null && source != null)
                        {
                            XLPivotTable pt;
                            switch (sourceType)
                            {
                                case XLPivotTableSourceType.Range:
                                    pt = ws.PivotTables.Add(pivotTableDefinition.Name, target, source) as XLPivotTable;
                                    break;

                                case XLPivotTableSourceType.Table:
                                    pt = ws.PivotTables.Add(pivotTableDefinition.Name, target, source as XLTable) as XLPivotTable;
                                    break;

                                default:
                                    throw new NotSupportedException($"Pivot table source type {sourceType} is not supported.");
                            }

                            if (!String.IsNullOrWhiteSpace(StringValue.ToString(pivotTableDefinition?.ColumnHeaderCaption ?? String.Empty)))
                                pt.SetColumnHeaderCaption(StringValue.ToString(pivotTableDefinition.ColumnHeaderCaption));

                            if (!String.IsNullOrWhiteSpace(StringValue.ToString(pivotTableDefinition?.RowHeaderCaption ?? String.Empty)))
                                pt.SetRowHeaderCaption(StringValue.ToString(pivotTableDefinition.RowHeaderCaption));

                            pt.RelId = worksheetPart.GetIdOfPart(pivotTablePart);
                            pt.CacheDefinitionRelId = pivotTablePart.GetIdOfPart(pivotTableCacheDefinitionPart);
                            pt.WorkbookCacheRelId = dSpreadsheet.WorkbookPart.GetIdOfPart(pivotTableCacheDefinitionPart);

                            if (pivotTableDefinition.MergeItem != null) pt.MergeAndCenterWithLabels = pivotTableDefinition.MergeItem.Value;
                            if (pivotTableDefinition.Indent != null) pt.RowLabelIndent = (int)pivotTableDefinition.Indent.Value;
                            if (pivotTableDefinition.PageOverThenDown != null) pt.FilterAreaOrder = pivotTableDefinition.PageOverThenDown.Value ? XLFilterAreaOrder.OverThenDown : XLFilterAreaOrder.DownThenOver;
                            if (pivotTableDefinition.PageWrap != null) pt.FilterFieldsPageWrap = (int)pivotTableDefinition.PageWrap.Value;
                            if (pivotTableDefinition.UseAutoFormatting != null) pt.AutofitColumns = pivotTableDefinition.UseAutoFormatting.Value;
                            if (pivotTableDefinition.PreserveFormatting != null) pt.PreserveCellFormatting = pivotTableDefinition.PreserveFormatting.Value;
                            if (pivotTableDefinition.RowGrandTotals != null) pt.ShowGrandTotalsRows = pivotTableDefinition.RowGrandTotals.Value;
                            if (pivotTableDefinition.ColumnGrandTotals != null) pt.ShowGrandTotalsColumns = pivotTableDefinition.ColumnGrandTotals.Value;
                            if (pivotTableDefinition.SubtotalHiddenItems != null) pt.FilteredItemsInSubtotals = pivotTableDefinition.SubtotalHiddenItems.Value;
                            if (pivotTableDefinition.MultipleFieldFilters != null) pt.AllowMultipleFilters = pivotTableDefinition.MultipleFieldFilters.Value;
                            if (pivotTableDefinition.CustomListSort != null) pt.UseCustomListsForSorting = pivotTableDefinition.CustomListSort.Value;
                            if (pivotTableDefinition.ShowDrill != null) pt.ShowExpandCollapseButtons = pivotTableDefinition.ShowDrill.Value;
                            if (pivotTableDefinition.ShowDataTips != null) pt.ShowContextualTooltips = pivotTableDefinition.ShowDataTips.Value;
                            if (pivotTableDefinition.ShowMemberPropertyTips != null) pt.ShowPropertiesInTooltips = pivotTableDefinition.ShowMemberPropertyTips.Value;
                            if (pivotTableDefinition.ShowHeaders != null) pt.DisplayCaptionsAndDropdowns = pivotTableDefinition.ShowHeaders.Value;
                            if (pivotTableDefinition.GridDropZones != null) pt.ClassicPivotTableLayout = pivotTableDefinition.GridDropZones.Value;
                            if (pivotTableDefinition.ShowEmptyRow != null) pt.ShowEmptyItemsOnRows = pivotTableDefinition.ShowEmptyRow.Value;
                            if (pivotTableDefinition.ShowEmptyColumn != null) pt.ShowEmptyItemsOnColumns = pivotTableDefinition.ShowEmptyColumn.Value;
                            if (pivotTableDefinition.ShowItems != null) pt.DisplayItemLabels = pivotTableDefinition.ShowItems.Value;
                            if (pivotTableDefinition.FieldListSortAscending != null) pt.SortFieldsAtoZ = pivotTableDefinition.FieldListSortAscending.Value;
                            if (pivotTableDefinition.PrintDrill != null) pt.PrintExpandCollapsedButtons = pivotTableDefinition.PrintDrill.Value;
                            if (pivotTableDefinition.ItemPrintTitles != null) pt.RepeatRowLabels = pivotTableDefinition.ItemPrintTitles.Value;
                            if (pivotTableDefinition.FieldPrintTitles != null) pt.PrintTitles = pivotTableDefinition.FieldPrintTitles.Value;
                            if (pivotTableDefinition.EnableDrill != null) pt.EnableShowDetails = pivotTableDefinition.EnableDrill.Value;
                            if (pivotTableCacheDefinitionPart.PivotCacheDefinition.SaveData != null) pt.SaveSourceData = pivotTableCacheDefinitionPart.PivotCacheDefinition.SaveData.Value;

                            if (pivotTableCacheDefinitionPart.PivotCacheDefinition.MissingItemsLimit != null)
                            {
                                if (pivotTableCacheDefinitionPart.PivotCacheDefinition.MissingItemsLimit == 0U)
                                    pt.ItemsToRetainPerField = XLItemsToRetain.None;
                                else if (pivotTableCacheDefinitionPart.PivotCacheDefinition.MissingItemsLimit == XLHelper.MaxRowNumber)
                                    pt.ItemsToRetainPerField = XLItemsToRetain.Max;
                            }

                            if (pivotTableDefinition.ShowMissing != null && pivotTableDefinition.MissingCaption != null)
                                pt.EmptyCellReplacement = pivotTableDefinition.MissingCaption.Value;

                            if (pivotTableDefinition.ShowError != null && pivotTableDefinition.ErrorCaption != null)
                                pt.ErrorValueReplacement = pivotTableDefinition.ErrorCaption.Value;

                            var pivotTableDefinitionExtensionList = pivotTableDefinition.GetFirstChild<PivotTableDefinitionExtensionList>();
                            var pivotTableDefinitionExtension = pivotTableDefinitionExtensionList?.GetFirstChild<PivotTableDefinitionExtension>();
                            var pivotTableDefinition2 = pivotTableDefinitionExtension?.GetFirstChild<DocumentFormat.OpenXml.Office2010.Excel.PivotTableDefinition>();
                            if (pivotTableDefinition2 != null)
                            {
                                if (pivotTableDefinition2.EnableEdit != null) pt.EnableCellEditing = pivotTableDefinition2.EnableEdit.Value;
                                if (pivotTableDefinition2.HideValuesRow != null) pt.ShowValuesRow = !pivotTableDefinition2.HideValuesRow.Value;
                            }

                            var pivotTableStyle = pivotTableDefinition.GetFirstChild<PivotTableStyle>();
                            if (pivotTableStyle != null)
                            {
                                pt.Theme = (XLPivotTableTheme)Enum.Parse(typeof(XLPivotTableTheme), pivotTableStyle.Name);
                                pt.ShowRowHeaders = pivotTableStyle.ShowRowHeaders;
                                pt.ShowColumnHeaders = pivotTableStyle.ShowColumnHeaders;
                                pt.ShowRowStripes = pivotTableStyle.ShowRowStripes;
                                pt.ShowColumnStripes = pivotTableStyle.ShowColumnStripes;
                            }

                            // Subtotal configuration
                            if (pivotTableDefinition.PivotFields.Cast<PivotField>().All(pf => pf.SubtotalTop != null && pf.SubtotalTop.HasValue && pf.SubtotalTop.Value))
                                pt.SetSubtotals(XLPivotSubtotals.AtTop);
                            else if (pivotTableDefinition.PivotFields.Cast<PivotField>().All(pf => pf.SubtotalTop != null && pf.SubtotalTop.HasValue && !pf.SubtotalTop.Value))
                                pt.SetSubtotals(XLPivotSubtotals.AtBottom);
                            else
                                pt.SetSubtotals(XLPivotSubtotals.DoNotShow);

                            // Row labels
                            if (pivotTableDefinition.RowFields != null)
                            {
                                foreach (var rf in pivotTableDefinition.RowFields.Cast<Field>())
                                {
                                    if (rf.Index < pivotTableDefinition.PivotFields.Count)
                                    {
                                        IXLPivotField pivotField = null;
                                        if (rf.Index.Value == -2)
                                            pivotField = pt.RowLabels.Add(XLConstants.PivotTableValuesSentinalLabel);
                                        else
                                        {
                                            var pf = pivotTableDefinition.PivotFields.ElementAt(rf.Index.Value) as PivotField;
                                            if (pf == null)
                                                continue;

                                            var cacheField = pivotTableCacheDefinitionPart.PivotCacheDefinition.CacheFields.ElementAt(rf.Index.Value) as CacheField;
                                            if (cacheField.Name != null)
                                                pivotField = pf.Name != null
                                                    ? pt.RowLabels.Add(cacheField.Name, pf.Name.Value)
                                                    : pt.RowLabels.Add(cacheField.Name.Value);
                                            else
                                                continue;

                                            if (pivotField != null)
                                            {
                                                LoadFieldOptions(pf, pivotField);
                                                LoadSubtotals(pf, pivotField);

                                                if (pf.SortType != null)
                                                {
                                                    pivotField.SetSort((XLPivotSortType)pf.SortType.Value);
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            // Column labels
                            if (pivotTableDefinition.ColumnFields != null)
                            {
                                foreach (var cf in pivotTableDefinition.ColumnFields.Cast<Field>())
                                {
                                    IXLPivotField pivotField = null;
                                    if (cf.Index.Value == -2)
                                        pivotField = pt.ColumnLabels.Add(XLConstants.PivotTableValuesSentinalLabel);
                                    else if (cf.Index < pivotTableDefinition.PivotFields.Count)
                                    {
                                        var pf = pivotTableDefinition.PivotFields.ElementAt(cf.Index.Value) as PivotField;
                                        if (pf == null)
                                            continue;

                                        var cacheField = pivotTableCacheDefinitionPart.PivotCacheDefinition.CacheFields.ElementAt(cf.Index.Value) as CacheField;
                                        if (cacheField.Name != null)
                                            pivotField = pf.Name != null
                                                ? pt.ColumnLabels.Add(cacheField.Name, pf.Name.Value)
                                                : pt.ColumnLabels.Add(cacheField.Name.Value);
                                        else
                                            continue;

                                        if (pivotField != null)
                                        {
                                            LoadFieldOptions(pf, pivotField);
                                            LoadSubtotals(pf, pivotField);

                                            if (pf.SortType != null)
                                            {
                                                pivotField.SetSort((XLPivotSortType)pf.SortType.Value);
                                            }
                                        }
                                    }
                                }
                            }

                            // Values
                            if (pivotTableDefinition.DataFields != null)
                            {
                                foreach (var df in pivotTableDefinition.DataFields.Cast<DataField>())
                                {
                                    IXLPivotValue pivotValue = null;
                                    if ((int)df.Field.Value == -2)
                                        pivotValue = pt.Values.Add(XLConstants.PivotTableValuesSentinalLabel);
                                    else if (df.Field.Value < pivotTableDefinition.PivotFields.Count)
                                    {
                                        var pf = pivotTableDefinition.PivotFields.ElementAt((int)df.Field.Value) as PivotField;
                                        if (pf == null)
                                            continue;

                                        var cacheField = pivotTableCacheDefinitionPart.PivotCacheDefinition.CacheFields.ElementAt((int)df.Field.Value) as CacheField;

                                        if (pf.Name != null)
                                            pivotValue = pt.Values.Add(pf.Name.Value, df.Name.Value);
                                        else if (cacheField.Name != null)
                                            pivotValue = pt.Values.Add(cacheField.Name.Value, df.Name.Value);
                                        else
                                            continue;

                                        if (df.NumberFormatId != null) pivotValue.NumberFormat.SetNumberFormatId((int)df.NumberFormatId.Value);
                                        if (df.Subtotal != null) pivotValue = pivotValue.SetSummaryFormula(df.Subtotal.Value.ToClosedXml());
                                        if (df.ShowDataAs != null)
                                        {
                                            var calculation = pivotValue.Calculation;
                                            calculation = df.ShowDataAs.Value.ToClosedXml();
                                            pivotValue = pivotValue.SetCalculation(calculation);
                                        }

                                        if (df.BaseField != null)
                                        {
                                            var col = pt.SourceRange.Column(df.BaseField.Value + 1);

                                            var items = col.CellsUsed()
                                                        .Select(c => c.Value)
                                                        .Skip(1) // Skip header column
                                                        .Distinct().ToList();

                                            pivotValue.BaseField = col.FirstCell().GetValue<string>();
                                            if (df.BaseItem != null) pivotValue.BaseItem = items[(int)df.BaseItem.Value].ToString();
                                        }
                                    }
                                }
                            }

                            // Filters
                            if (pivotTableDefinition.PageFields != null)
                            {
                                foreach (var pageField in pivotTableDefinition.PageFields.Cast<PageField>())
                                {
                                    var pf = pivotTableDefinition.PivotFields.ElementAt((int)pageField.Field.Value) as PivotField;
                                    if (pf == null)
                                        continue;

                                    var cacheField = pivotTableCacheDefinitionPart.PivotCacheDefinition.CacheFields.ElementAt((int)pageField.Field.Value) as CacheField;

                                    var filterName = pf.Name?.Value ?? cacheField.Name?.Value;

                                    IXLPivotField rf;
                                    if (pageField.Name?.Value != null)
                                        rf = pt.ReportFilters.Add(filterName, pageField.Name.Value);
                                    else
                                        rf = pt.ReportFilters.Add(filterName);

                                    if ((pageField.Item?.HasValue ?? false)
                                        && pf.Items.Any() && cacheField.SharedItems.Any())
                                    {
                                        var item = pf.Items.ElementAt(Convert.ToInt32(pageField.Item.Value)) as Item;
                                        if (item == null)
                                            continue;

                                        var sharedItem = cacheField.SharedItems.ElementAt(Convert.ToInt32((uint)item.Index));
                                        var numberItem = sharedItem as NumberItem;
                                        var stringItem = sharedItem as StringItem;
                                        var dateTimeItem = sharedItem as DateTimeItem;

                                        if (numberItem != null)
                                            rf.AddSelectedValue(Convert.ToDouble(numberItem.Val.Value));
                                        else if (dateTimeItem != null)
                                            rf.AddSelectedValue(Convert.ToDateTime(dateTimeItem.Val.Value));
                                        else if (stringItem != null)
                                            rf.AddSelectedValue(stringItem.Val.Value);
                                        else
                                            throw new NotImplementedException();
                                    }
                                    else if (OpenXmlHelper.GetBooleanValueAsBool(pf.MultipleItemSelectionAllowed, false))
                                    {
                                        foreach (var item in pf.Items.Cast<Item>())
                                        {
                                            if (item.Hidden == null || !BooleanValue.ToBoolean(item.Hidden))
                                            {
                                                var sharedItem = cacheField.SharedItems.ElementAt(Convert.ToInt32((uint)item.Index));
                                                var numberItem = sharedItem as NumberItem;
                                                var stringItem = sharedItem as StringItem;
                                                var dateTimeItem = sharedItem as DateTimeItem;

                                                if (numberItem != null)
                                                    rf.AddSelectedValue(Convert.ToDouble(numberItem.Val.Value));
                                                else if (dateTimeItem != null)
                                                    rf.AddSelectedValue(Convert.ToDateTime(dateTimeItem.Val.Value));
                                                else if (stringItem != null)
                                                    rf.AddSelectedValue(stringItem.Val.Value);
                                                else
                                                    throw new NotImplementedException();
                                            }
                                        }
                                    }
                                }

                                pt.TargetCell = pt.TargetCell.CellAbove(pt.ReportFilters.Count() + 1);
                            }
                        }
                    }
                }
            }

            #endregion Pivot tables
        }

        private static void LoadFieldOptions(PivotField pf, IXLPivotField pivotField)
        {
            if (pf.SubtotalCaption != null) pivotField.SubtotalCaption = pf.SubtotalCaption;
            if (pf.IncludeNewItemsInFilter != null) pivotField.IncludeNewItemsInFilter = pf.IncludeNewItemsInFilter.Value;
            if (pf.Outline != null) pivotField.Outline = pf.Outline.Value;
            if (pf.Compact != null) pivotField.Compact = pf.Compact.Value;
            if (pf.InsertBlankRow != null) pivotField.InsertBlankLines = pf.InsertBlankRow.Value;
            if (pf.ShowAll != null) pivotField.ShowBlankItems = pf.ShowAll.Value;
            if (pf.InsertPageBreak != null) pivotField.InsertPageBreaks = pf.InsertPageBreak.Value;
            if (pf.SubtotalTop != null) pivotField.SubtotalsAtTop = pf.SubtotalTop.Value;
            if (pf.AllDrilled != null) pivotField.Collapsed = !pf.AllDrilled.Value;

            var pivotFieldExtensionList = pf.GetFirstChild<PivotFieldExtensionList>();
            var pivotFieldExtension = pivotFieldExtensionList?.GetFirstChild<PivotFieldExtension>();
            var field2010 = pivotFieldExtension?.GetFirstChild<DocumentFormat.OpenXml.Office2010.Excel.PivotField>();
            if (field2010?.FillDownLabels != null) pivotField.RepeatItemLabels = field2010.FillDownLabels.Value;
        }

        private static void LoadSubtotals(PivotField pf, IXLPivotField pivotField)
        {
            if (pf.AverageSubTotal != null)
                pivotField.AddSubtotal(XLSubtotalFunction.Average);
            if (pf.CountASubtotal != null)
                pivotField.AddSubtotal(XLSubtotalFunction.Count);
            if (pf.CountSubtotal != null)
                pivotField.AddSubtotal(XLSubtotalFunction.CountNumbers);
            if (pf.MaxSubtotal != null)
                pivotField.AddSubtotal(XLSubtotalFunction.Maximum);
            if (pf.MinSubtotal != null)
                pivotField.AddSubtotal(XLSubtotalFunction.Minimum);
            if (pf.ApplyStandardDeviationPInSubtotal != null)
                pivotField.AddSubtotal(XLSubtotalFunction.PopulationStandardDeviation);
            if (pf.ApplyVariancePInSubtotal != null)
                pivotField.AddSubtotal(XLSubtotalFunction.PopulationVariance);
            if (pf.ApplyProductInSubtotal != null)
                pivotField.AddSubtotal(XLSubtotalFunction.Product);
            if (pf.ApplyStandardDeviationInSubtotal != null)
                pivotField.AddSubtotal(XLSubtotalFunction.StandardDeviation);
            if (pf.SumSubtotal != null)
                pivotField.AddSubtotal(XLSubtotalFunction.Sum);
            if (pf.ApplyVarianceInSubtotal != null)
                pivotField.AddSubtotal(XLSubtotalFunction.Variance);

            if (pf.Items?.Any() ?? false)
            {
                var items = pf.Items.OfType<Item>().Where(i => i.Index != null && i.Index.HasValue);
                if (!items.Any(i => i.HideDetails == null || BooleanValue.ToBoolean(i.HideDetails)))
                    pivotField.SetCollapsed();
            }
        }

        private void LoadDrawings(WorksheetPart wsPart, IXLWorksheet ws)
        {
            if (wsPart.DrawingsPart != null)
            {
                var drawingsPart = wsPart.DrawingsPart;

                foreach (var anchor in drawingsPart.WorksheetDrawing.ChildElements)
                {
                    var imgId = GetImageRelIdFromAnchor(anchor);

                    //If imgId is null, we're probably dealing with a TextBox (or another shape) instead of a picture
                    if (imgId == null) continue;

                    var imagePart = drawingsPart.GetPartById(imgId);
                    using (var stream = imagePart.GetStream())
                    using (var ms = new MemoryStream())
                    {
                        stream.CopyTo(ms);
                        var vsdp = GetPropertiesFromAnchor(anchor);

                        var picture = (ws as XLWorksheet).AddPicture(ms, vsdp.Name, Convert.ToInt32(vsdp.Id.Value)) as XLPicture;
                        picture.RelId = imgId;

                        Xdr.ShapeProperties spPr = anchor.Descendants<Xdr.ShapeProperties>().First();
                        picture.Placement = XLPicturePlacement.FreeFloating;

                        if (spPr?.Transform2D?.Extents?.Cx.HasValue ?? false)
                            picture.Width = ConvertFromEnglishMetricUnits(spPr.Transform2D.Extents.Cx, GraphicsUtils.Graphics.DpiX);

                        if (spPr?.Transform2D?.Extents?.Cy.HasValue ?? false)
                            picture.Height = ConvertFromEnglishMetricUnits(spPr.Transform2D.Extents.Cy, GraphicsUtils.Graphics.DpiY);

                        if (anchor is Xdr.AbsoluteAnchor)
                        {
                            var absoluteAnchor = anchor as Xdr.AbsoluteAnchor;
                            picture.MoveTo(
                                ConvertFromEnglishMetricUnits(absoluteAnchor.Position.X.Value, GraphicsUtils.Graphics.DpiX),
                                ConvertFromEnglishMetricUnits(absoluteAnchor.Position.Y.Value, GraphicsUtils.Graphics.DpiY)
                            );
                        }
                        else if (anchor is Xdr.OneCellAnchor)
                        {
                            var oneCellAnchor = anchor as Xdr.OneCellAnchor;
                            var from = LoadMarker(ws, oneCellAnchor.FromMarker);
                            picture.MoveTo(from.Address, from.Offset);
                        }
                        else if (anchor is Xdr.TwoCellAnchor)
                        {
                            var twoCellAnchor = anchor as Xdr.TwoCellAnchor;
                            var from = LoadMarker(ws, twoCellAnchor.FromMarker);
                            var to = LoadMarker(ws, twoCellAnchor.ToMarker);

                            if (twoCellAnchor.EditAs == null || !twoCellAnchor.EditAs.HasValue || twoCellAnchor.EditAs.Value == Xdr.EditAsValues.TwoCell)
                            {
                                picture.MoveTo(from.Address, from.Offset, to.Address, to.Offset);
                            }
                            else if (twoCellAnchor.EditAs.Value == Xdr.EditAsValues.Absolute)
                            {
                                var shapeProperties = twoCellAnchor.Descendants<Xdr.ShapeProperties>().FirstOrDefault();
                                if (shapeProperties != null)
                                {
                                    picture.MoveTo(
                                        ConvertFromEnglishMetricUnits(spPr.Transform2D.Offset.X, GraphicsUtils.Graphics.DpiX),
                                        ConvertFromEnglishMetricUnits(spPr.Transform2D.Offset.Y, GraphicsUtils.Graphics.DpiY)
                                    );
                                }
                            }
                            else if (twoCellAnchor.EditAs.Value == Xdr.EditAsValues.OneCell)
                            {
                                picture.MoveTo(from.Address, from.Offset);
                            }
                        }
                    }
                }
            }
        }

        private static Int32 ConvertFromEnglishMetricUnits(long emu, float resolution)
        {
            return Convert.ToInt32(emu * resolution / 914400);
        }

        private static IXLMarker LoadMarker(IXLWorksheet ws, Xdr.MarkerType marker)
        {
            var row = Math.Min(XLHelper.MaxRowNumber, Math.Max(1, Convert.ToInt32(marker.RowId.InnerText) + 1));
            var column = Math.Min(XLHelper.MaxColumnNumber, Math.Max(1, Convert.ToInt32(marker.ColumnId.InnerText) + 1));
            return new XLMarker(
                ws.Cell(row, column).Address,
                new Point(
                    ConvertFromEnglishMetricUnits(Convert.ToInt32(marker.ColumnOffset.InnerText), GraphicsUtils.Graphics.DpiX),
                    ConvertFromEnglishMetricUnits(Convert.ToInt32(marker.RowOffset.InnerText), GraphicsUtils.Graphics.DpiY)
                )
            );
        }

        #region Comment Helpers

        private XDocument GetCommentVmlFile(WorksheetPart wsPart)
        {
            XDocument xdoc = null;

            foreach (var vmlPart in wsPart.VmlDrawingParts)
            {
                xdoc = XDocumentExtensions.Load(vmlPart.GetStream(FileMode.Open));

                //Probe for comments
                if (xdoc?.Root == null) continue;
                var shape = GetCommentShape(xdoc);
                if (shape != null) break;
            }

            if (xdoc == null) throw new ArgumentException("Could not load comments file");
            return xdoc;
        }

        private static XElement GetCommentShape(XDocument xdoc)
        {
            var xml = xdoc.Root.Element("xml");

            XElement shape;
            if (xml != null)
                shape =
                    xml.Elements().FirstOrDefault(e => (string)e.Attribute("type") == XLConstants.Comment.ShapeTypeId);
            else
                shape = xdoc.Root.Elements().FirstOrDefault(e =>
                                                            (string)e.Attribute("type") ==
                                                            XLConstants.Comment.ShapeTypeId ||
                                                            (string)e.Attribute("type") ==
                                                            XLConstants.Comment.AlternateShapeTypeId);
            return shape;
        }

        #endregion Comment Helpers

        private String GetTableColumnName(string name)
        {
            return name.Replace("_x000a_", Environment.NewLine).Replace("_x005f_x000a_", "_x000a_");
        }

        // This may be part of XLHelper or XLColor
        // Leaving it here for now. Can't decide what to call it and where to put it.
        private XLColor ExtractColor(String color)
        {
            if (color.IndexOf("[") >= 0)
            {
                int start = color.IndexOf("[") + 1;
                int end = color.IndexOf("]", start);
                return XLColor.FromIndex(Int32.Parse(color.Substring(start, end - start)));
            }
            else
            {
                return XLColor.FromHtml(color);
            }
        }

        private void LoadColorsAndLines<T>(IXLDrawing<T> drawing, XElement shape)
        {
            var strokeColor = shape.Attribute("strokecolor");
            if (strokeColor != null) drawing.Style.ColorsAndLines.LineColor = ExtractColor(strokeColor.Value);

            var strokeWeight = shape.Attribute("strokeweight");
            if (strokeWeight != null)
                drawing.Style.ColorsAndLines.LineWeight = GetPtValue(strokeWeight.Value);

            var fillColor = shape.Attribute("fillcolor");
            if (fillColor != null && !fillColor.Value.ToLower().Contains("infobackground")) drawing.Style.ColorsAndLines.FillColor = ExtractColor(fillColor.Value);

            var fill = shape.Elements().FirstOrDefault(e => e.Name.LocalName == "fill");
            if (fill != null)
            {
                var opacity = fill.Attribute("opacity");
                if (opacity != null)
                {
                    String opacityVal = opacity.Value;
                    if (opacityVal.EndsWith("f"))
                        drawing.Style.ColorsAndLines.FillTransparency =
                            Double.Parse(opacityVal.Substring(0, opacityVal.Length - 1), CultureInfo.InvariantCulture) / 65536.0;
                    else
                        drawing.Style.ColorsAndLines.FillTransparency = Double.Parse(opacityVal, CultureInfo.InvariantCulture);
                }
            }

            var stroke = shape.Elements().FirstOrDefault(e => e.Name.LocalName == "stroke");
            if (stroke != null)
            {
                var opacity = stroke.Attribute("opacity");
                if (opacity != null)
                {
                    String opacityVal = opacity.Value;
                    if (opacityVal.EndsWith("f"))
                        drawing.Style.ColorsAndLines.LineTransparency =
                            Double.Parse(opacityVal.Substring(0, opacityVal.Length - 1), CultureInfo.InvariantCulture) / 65536.0;
                    else
                        drawing.Style.ColorsAndLines.LineTransparency = Double.Parse(opacityVal, CultureInfo.InvariantCulture);
                }

                var dashStyle = stroke.Attribute("dashstyle");
                if (dashStyle != null)
                {
                    String dashStyleVal = dashStyle.Value.ToLower();
                    if (dashStyleVal == "1 1" || dashStyleVal == "shortdot")
                    {
                        var endCap = stroke.Attribute("endcap");
                        if (endCap != null && endCap.Value == "round")
                            drawing.Style.ColorsAndLines.LineDash = XLDashStyle.RoundDot;
                        else
                            drawing.Style.ColorsAndLines.LineDash = XLDashStyle.SquareDot;
                    }
                    else
                    {
                        switch (dashStyleVal)
                        {
                            case "dash": drawing.Style.ColorsAndLines.LineDash = XLDashStyle.Dash; break;
                            case "dashdot": drawing.Style.ColorsAndLines.LineDash = XLDashStyle.DashDot; break;
                            case "longdash": drawing.Style.ColorsAndLines.LineDash = XLDashStyle.LongDash; break;
                            case "longdashdot": drawing.Style.ColorsAndLines.LineDash = XLDashStyle.LongDashDot; break;
                            case "longdashdotdot": drawing.Style.ColorsAndLines.LineDash = XLDashStyle.LongDashDotDot; break;
                        }
                    }
                }

                var lineStyle = stroke.Attribute("linestyle");
                if (lineStyle != null)
                {
                    String lineStyleVal = lineStyle.Value.ToLower();
                    switch (lineStyleVal)
                    {
                        case "single": drawing.Style.ColorsAndLines.LineStyle = XLLineStyle.Single; break;
                        case "thickbetweenthin": drawing.Style.ColorsAndLines.LineStyle = XLLineStyle.ThickBetweenThin; break;
                        case "thickthin": drawing.Style.ColorsAndLines.LineStyle = XLLineStyle.ThickThin; break;
                        case "thinthick": drawing.Style.ColorsAndLines.LineStyle = XLLineStyle.ThinThick; break;
                        case "thinthin": drawing.Style.ColorsAndLines.LineStyle = XLLineStyle.ThinThin; break;
                    }
                }
            }
        }

        private void LoadTextBox<T>(IXLDrawing<T> xlDrawing, XElement textBox)
        {
            var attStyle = textBox.Attribute("style");
            if (attStyle != null) LoadTextBoxStyle<T>(xlDrawing, attStyle);

            var attInset = textBox.Attribute("inset");
            if (attInset != null) LoadTextBoxInset<T>(xlDrawing, attInset);
        }

        private void LoadTextBoxInset<T>(IXLDrawing<T> xlDrawing, XAttribute attInset)
        {
            var split = attInset.Value.Split(',');
            xlDrawing.Style.Margins.Left = GetInsetValue(split[0]);
            xlDrawing.Style.Margins.Top = GetInsetValue(split[1]);
            xlDrawing.Style.Margins.Right = GetInsetValue(split[2]);
            xlDrawing.Style.Margins.Bottom = GetInsetValue(split[3]);
        }

        private double GetInsetValue(string value)
        {
            String v = value.Trim();
            if (v.EndsWith("pt"))
                return Double.Parse(v.Substring(0, v.Length - 2), CultureInfo.InvariantCulture) / 72.0;
            else
                return Double.Parse(v.Substring(0, v.Length - 2), CultureInfo.InvariantCulture);
        }

        private static void LoadTextBoxStyle<T>(IXLDrawing<T> xlDrawing, XAttribute attStyle)
        {
            var style = attStyle.Value;
            var attributes = style.Split(';');
            foreach (String pair in attributes)
            {
                var split = pair.Split(':');
                if (split.Length != 2) continue;

                var attribute = split[0].Trim().ToLower();
                var value = split[1].Trim();
                Boolean isVertical = false;
                switch (attribute)
                {
                    case "mso-fit-shape-to-text": xlDrawing.Style.Size.SetAutomaticSize(value.Equals("t")); break;
                    case "mso-layout-flow-alt":
                        if (value.Equals("bottom-to-top")) xlDrawing.Style.Alignment.SetOrientation(XLDrawingTextOrientation.BottomToTop);
                        else if (value.Equals("top-to-bottom")) xlDrawing.Style.Alignment.SetOrientation(XLDrawingTextOrientation.Vertical);
                        break;

                    case "layout-flow": isVertical = value.Equals("vertical"); break;
                    case "mso-direction-alt": if (value == "auto") xlDrawing.Style.Alignment.Direction = XLDrawingTextDirection.Context; break;
                    case "direction": if (value == "RTL") xlDrawing.Style.Alignment.Direction = XLDrawingTextDirection.RightToLeft; break;
                }
                if (isVertical && xlDrawing.Style.Alignment.Orientation == XLDrawingTextOrientation.LeftToRight)
                    xlDrawing.Style.Alignment.Orientation = XLDrawingTextOrientation.TopToBottom;
            }
        }

        private void LoadClientData<T>(IXLDrawing<T> drawing, XElement clientData)
        {
            var anchor = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "Anchor");
            if (anchor != null) LoadClientDataAnchor<T>(drawing, anchor);

            LoadDrawingPositioning<T>(drawing, clientData);
            LoadDrawingProtection<T>(drawing, clientData);

            var visible = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "Visible");
            drawing.Visible = visible != null && visible.Value.ToLower().StartsWith("t");

            LoadDrawingHAlignment<T>(drawing, clientData);
            LoadDrawingVAlignment<T>(drawing, clientData);
        }

        private void LoadDrawingHAlignment<T>(IXLDrawing<T> drawing, XElement clientData)
        {
            var textHAlign = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "TextHAlign");
            if (textHAlign != null)
                drawing.Style.Alignment.Horizontal = (XLDrawingHorizontalAlignment)Enum.Parse(typeof(XLDrawingHorizontalAlignment), textHAlign.Value.ToProper());
        }

        private void LoadDrawingVAlignment<T>(IXLDrawing<T> drawing, XElement clientData)
        {
            var textVAlign = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "TextVAlign");
            if (textVAlign != null)
                drawing.Style.Alignment.Vertical = (XLDrawingVerticalAlignment)Enum.Parse(typeof(XLDrawingVerticalAlignment), textVAlign.Value.ToProper());
        }

        private void LoadDrawingProtection<T>(IXLDrawing<T> drawing, XElement clientData)
        {
            var lockedElement = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "Locked");
            var lockTextElement = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "LockText");
            Boolean locked = lockedElement != null && lockedElement.Value.ToLower() == "true";
            Boolean lockText = lockTextElement != null && lockTextElement.Value.ToLower() == "true";
            drawing.Style.Protection.Locked = locked;
            drawing.Style.Protection.LockText = lockText;
        }

        private static void LoadDrawingPositioning<T>(IXLDrawing<T> drawing, XElement clientData)
        {
            var moveWithCellsElement = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "MoveWithCells");
            var sizeWithCellsElement = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "SizeWithCells");
            Boolean moveWithCells = !(moveWithCellsElement != null && moveWithCellsElement.Value.ToLower() == "true");
            Boolean sizeWithCells = !(sizeWithCellsElement != null && sizeWithCellsElement.Value.ToLower() == "true");
            if (moveWithCells && !sizeWithCells)
                drawing.Style.Properties.Positioning = XLDrawingAnchor.MoveWithCells;
            else if (moveWithCells && sizeWithCells)
                drawing.Style.Properties.Positioning = XLDrawingAnchor.MoveAndSizeWithCells;
            else
                drawing.Style.Properties.Positioning = XLDrawingAnchor.Absolute;
        }

        private static void LoadClientDataAnchor<T>(IXLDrawing<T> drawing, XElement anchor)
        {
            var location = anchor.Value.Split(',');
            drawing.Position.Column = int.Parse(location[0]) + 1;
            drawing.Position.ColumnOffset = Double.Parse(location[1], CultureInfo.InvariantCulture) / 7.2;
            drawing.Position.Row = int.Parse(location[2]) + 1;
            drawing.Position.RowOffset = Double.Parse(location[3], CultureInfo.InvariantCulture);
        }

        private void LoadShapeProperties<T>(IXLDrawing<T> xlDrawing, XElement shape)
        {
            var attStyle = shape.Attribute("style");
            if (attStyle == null) return;

            var style = attStyle.Value;
            var attributes = style.Split(';');
            foreach (String pair in attributes)
            {
                var split = pair.Split(':');
                if (split.Length != 2) continue;

                var attribute = split[0].Trim().ToLower();
                var value = split[1].Trim();

                switch (attribute)
                {
                    case "visibility": xlDrawing.Visible = value.ToLower().Equals("visible"); break;
                    case "width": xlDrawing.Style.Size.Width = GetPtValue(value) / 7.5; break;
                    case "height": xlDrawing.Style.Size.Height = GetPtValue(value); break;
                    case "z-index": xlDrawing.ZOrder = Int32.Parse(value); break;
                }
            }
        }

        private readonly Dictionary<string, double> knownUnits = new Dictionary<string, double>
        {
            {"pt", 1.0},
            {"in", 72.0},
            {"mm", 72.0/25.4}
        };

        private double GetPtValue(string value)
        {
            var knownUnit = knownUnits.FirstOrDefault(ku => value.Contains(ku.Key));

            if (knownUnit.Key == null)
                return Double.Parse(value);

            return Double.Parse(value.Replace(knownUnit.Key, String.Empty), CultureInfo.InvariantCulture) * knownUnit.Value;
        }

        private void LoadDefinedNames(Workbook workbook)
        {
            if (workbook.DefinedNames == null) return;

            foreach (var definedName in workbook.DefinedNames.OfType<DefinedName>())
            {
                var name = definedName.Name;
                var visible = true;
                if (definedName.Hidden != null) visible = !BooleanValue.ToBoolean(definedName.Hidden);
                if (name == "_xlnm.Print_Area")
                {
                    var fixedNames = validateDefinedNames(definedName.Text.Split(','));
                    foreach (string area in fixedNames)
                    {
                        if (area.Contains("["))
                        {
                            var ws = Worksheets.FirstOrDefault(w => (w as XLWorksheet).SheetId == definedName.LocalSheetId + 1);
                            if (ws != null)
                            {
                                ws.PageSetup.PrintAreas.Add(area);
                            }
                        }
                        else
                        {
                            ParseReference(area, out String sheetName, out String sheetArea);
                            if (!(sheetArea.Equals("#REF") || sheetArea.EndsWith("#REF!") || sheetArea.Length == 0 || sheetName.Length == 0))
                                WorksheetsInternal.Worksheet(sheetName).PageSetup.PrintAreas.Add(sheetArea);
                        }
                    }
                }
                else if (name == "_xlnm.Print_Titles")
                {
                    LoadPrintTitles(definedName);
                }
                else
                {
                    string text = definedName.Text;

                    if (!(text.Equals("#REF") || text.EndsWith("#REF!")))
                    {
                        var localSheetId = definedName.LocalSheetId;
                        var comment = definedName.Comment;
                        if (localSheetId == null)
                        {
                            if (NamedRanges.All(nr => nr.Name != name))
                                (NamedRanges as XLNamedRanges).Add(name, text, comment, true).Visible = visible;
                        }
                        else
                        {
                            if (Worksheet(Int32.Parse(localSheetId) + 1).NamedRanges.All(nr => nr.Name != name))
                                (Worksheet(Int32.Parse(localSheetId) + 1).NamedRanges as XLNamedRanges).Add(name, text, comment, true).Visible = visible;
                        }
                    }
                }
            }
        }

        private static Regex definedNameRegex = new Regex(@"\A'.*'!.*\z", RegexOptions.Compiled);

        private IEnumerable<String> validateDefinedNames(IEnumerable<String> definedNames)
        {
            var sb = new StringBuilder();
            foreach (string testName in definedNames)
            {
                if (sb.Length > 0)
                    sb.Append(',');

                sb.Append(testName);

                Match matchedValidPattern = definedNameRegex.Match(sb.ToString());
                if (matchedValidPattern.Success)
                {
                    yield return sb.ToString();
                    sb = new StringBuilder();
                }
            }

            if (sb.Length > 0)
                yield return sb.ToString();
        }

        private void LoadPrintTitles(DefinedName definedName)
        {
            var areas = validateDefinedNames(definedName.Text.Split(','));
            foreach (var item in areas)
            {
                if (this.Range(item) != null)
                    SetColumnsOrRowsToRepeat(item);
            }
        }

        private void SetColumnsOrRowsToRepeat(string area)
        {
            ParseReference(area, out String sheetName, out String sheetArea);
            if (sheetArea.Equals("#REF")) return;
            if (IsColReference(sheetArea))
                WorksheetsInternal.Worksheet(sheetName).PageSetup.SetColumnsToRepeatAtLeft(sheetArea);
            if (IsRowReference(sheetArea))
                WorksheetsInternal.Worksheet(sheetName).PageSetup.SetRowsToRepeatAtTop(sheetArea);
        }

        // either $A:$X => true or $1:$99 => false
        private static bool IsColReference(string sheetArea)
        {
            char c = sheetArea[0] == '$' ? sheetArea[1] : sheetArea[0];
            return char.IsLetter(c);
        }

        private static bool IsRowReference(string sheetArea)
        {
            char c = sheetArea[0] == '$' ? sheetArea[1] : sheetArea[0];
            return char.IsNumber(c);
        }

        private static void ParseReference(string item, out string sheetName, out string sheetArea)
        {
            var sections = item.Trim().Split('!');
            if (sections.Count() == 1)
            {
                sheetName = string.Empty;
                sheetArea = item;
            }
            else
            {
                sheetName = sections[0].UnescapeSheetName();
                sheetArea = sections[1];
            }
        }

        private Int32 lastCell;

        private void LoadCells(SharedStringItem[] sharedStrings, Stylesheet s, NumberingFormats numberingFormats,
                               Fills fills, Borders borders, Fonts fonts, Dictionary<uint, string> sharedFormulasR1C1,
                               XLWorksheet ws, Dictionary<Int32, IXLStyle> styleList, Cell cell, Int32 rowIndex)
        {
            Int32 styleIndex = cell.StyleIndex != null ? Int32.Parse(cell.StyleIndex.InnerText) : 0;

            String cellReference = cell.CellReference == null
                                       ? XLHelper.GetColumnLetterFromNumber(++lastCell) + rowIndex
                                       : cell.CellReference.Value;
            var xlCell = ws.CellFast(cellReference);

            if (styleList.ContainsKey(styleIndex))
            {
                xlCell.Style = styleList[styleIndex];
            }
            else
            {
                ApplyStyle(xlCell, styleIndex, s, fills, borders, fonts, numberingFormats);
            }

            if (cell.CellFormula?.SharedIndex != null && cell.CellFormula?.Reference != null)
            {
                String formula;
                if (cell.CellFormula.FormulaType != null && cell.CellFormula.FormulaType == CellFormulaValues.Array)
                    formula = "{" + cell.CellFormula.Text + "}";
                else
                    formula = cell.CellFormula.Text;

                // Parent cell of shared formulas
                // Child cells will use this shared index to set its R1C1 style formula
                xlCell.FormulaReference = ws.Range(cell.CellFormula.Reference.Value).RangeAddress;

                xlCell.FormulaA1 = formula;
                sharedFormulasR1C1.Add(cell.CellFormula.SharedIndex.Value, xlCell.FormulaR1C1);

                if (cell.DataType != null)
                {
                    switch (cell.DataType.Value)
                    {
                        case CellValues.Boolean:
                            xlCell.SetDataTypeFast(XLDataType.Boolean);
                            break;

                        case CellValues.Number:
                            xlCell.SetDataTypeFast(XLDataType.Number);
                            break;

                        case CellValues.Date:
                            xlCell.SetDataTypeFast(XLDataType.DateTime);
                            break;

                        case CellValues.InlineString:
                        case CellValues.SharedString:
                            xlCell.SetDataTypeFast(XLDataType.Text);
                            break;
                    }
                }

                if (cell.CellValue != null)
                {
#pragma warning disable 618
                    xlCell.ValueCached = cell.CellValue.Text;
#pragma warning restore 618
                    xlCell.SetInternalCellValueString(cell.CellValue.Text);
                }
            }
            else if (cell.CellFormula != null)
            {
                if (cell.CellFormula.SharedIndex != null)
                    xlCell.FormulaR1C1 = sharedFormulasR1C1[cell.CellFormula.SharedIndex.Value];
                else if (!String.IsNullOrWhiteSpace(cell.CellFormula.Text))
                {
                    String formula;
                    if (cell.CellFormula.FormulaType != null && cell.CellFormula.FormulaType == CellFormulaValues.Array)
                        formula = "{" + cell.CellFormula.Text + "}";
                    else
                        formula = cell.CellFormula.Text;

                    xlCell.FormulaA1 = formula;
                }

                if (cell.CellFormula.Reference != null)
                {
                    foreach (var childCell in ws.Range(cell.CellFormula.Reference.Value).Cells(c => c.FormulaReference == null || !c.HasFormula))
                    {
                        if (childCell.FormulaReference == null)
                            childCell.FormulaReference = ws.Range(cell.CellFormula.Reference.Value).RangeAddress;

                        if (!childCell.HasFormula)
                            childCell.FormulaA1 = xlCell.FormulaA1;
                    }
                }

                if (cell.DataType != null)
                {
                    switch (cell.DataType.Value)
                    {
                        case CellValues.Boolean:
                            xlCell.SetDataTypeFast(XLDataType.Boolean);
                            break;

                        case CellValues.Number:
                            xlCell.SetDataTypeFast(XLDataType.Number);
                            break;

                        case CellValues.Date:
                            xlCell.SetDataTypeFast(XLDataType.DateTime);
                            break;

                        case CellValues.InlineString:
                        case CellValues.SharedString:
                            xlCell.SetDataTypeFast(XLDataType.Text);
                            break;
                    }
                }

                if (cell.CellValue != null)
                {
#pragma warning disable 618
                    xlCell.ValueCached = cell.CellValue.Text;
#pragma warning restore 618
                    xlCell.SetInternalCellValueString(cell.CellValue.Text);
                }
            }
            else if (cell.DataType != null)
            {
                if (cell.DataType == CellValues.InlineString)
                {
                    xlCell.SetDataTypeFast(XLDataType.Text);
                    xlCell.ShareString = false;

                    if (cell.InlineString != null)
                    {
                        if (cell.InlineString.Text != null)
                            xlCell.SetInternalCellValueString(cell.InlineString.Text.Text.FixNewLines());
                        else
                            ParseCellValue(cell.InlineString, xlCell);
                    }
                    else
                        xlCell.SetInternalCellValueString(String.Empty);
                }
                else if (cell.DataType == CellValues.SharedString)
                {
                    xlCell.SetDataTypeFast(XLDataType.Text);

                    if (cell.CellValue != null && !String.IsNullOrWhiteSpace(cell.CellValue.Text))
                    {
                        var sharedString = sharedStrings[Int32.Parse(cell.CellValue.Text, XLHelper.NumberStyle, XLHelper.ParseCulture)];
                        ParseCellValue(sharedString, xlCell);
                    }
                    else
                        xlCell.SetInternalCellValueString(String.Empty);
                }
                else if (cell.DataType == CellValues.Date)
                {
                    xlCell.SetDataTypeFast(XLDataType.DateTime);

                    if (cell.CellValue != null && !String.IsNullOrWhiteSpace(cell.CellValue.Text))
                        xlCell.SetInternalCellValueString(Double.Parse(cell.CellValue.Text, XLHelper.NumberStyle, XLHelper.ParseCulture).ToInvariantString());
                }
                else if (cell.DataType == CellValues.Boolean)
                {
                    xlCell.SetDataTypeFast(XLDataType.Boolean);
                    if (cell.CellValue != null)
                        xlCell.SetInternalCellValueString(cell.CellValue.Text);
                }
                else if (cell.DataType == CellValues.Number)
                {
                    if (s == null)
                        xlCell.SetDataTypeFast(XLDataType.Number);
                    else
                        xlCell.DataType = GetDataTypeFromCell(xlCell.Style.NumberFormat);

                    if (cell.CellValue != null && !String.IsNullOrWhiteSpace(cell.CellValue.Text))
                        xlCell.SetInternalCellValueString(Double.Parse(cell.CellValue.Text, XLHelper.NumberStyle, XLHelper.ParseCulture).ToInvariantString());
                }
            }
            else if (cell.CellValue != null)
            {
                if (s == null)
                {
                    xlCell.SetDataTypeFast(XLDataType.Number);
                }
                else
                {
                    xlCell.DataType = GetDataTypeFromCell(xlCell.Style.NumberFormat);
                    var numberFormatId = ((CellFormat)(s.CellFormats).ElementAt(styleIndex)).NumberFormatId;
                    if (!String.IsNullOrWhiteSpace(cell.CellValue.Text))
                        xlCell.SetInternalCellValueString(Double.Parse(cell.CellValue.Text, CultureInfo.InvariantCulture).ToInvariantString());

                    if (s.NumberingFormats != null &&
                        s.NumberingFormats.Any(nf => ((NumberingFormat)nf).NumberFormatId.Value == numberFormatId))
                    {
                        xlCell.Style.NumberFormat.Format =
                            ((NumberingFormat)s.NumberingFormats
                                                .First(
                                                    nf => ((NumberingFormat)nf).NumberFormatId.Value == numberFormatId)
                            ).FormatCode.Value;
                    }
                    else
                        xlCell.Style.NumberFormat.NumberFormatId = Int32.Parse(numberFormatId);
                }
            }

            if (Use1904DateSystem && xlCell.DataType == XLDataType.DateTime)
            {
                // Internally ClosedXML stores cells as standard 1900-based style
                // so if a workbook is in 1904-format, we do that adjustment here and when saving.
                xlCell.SetValue(xlCell.GetDateTime().AddDays(1462));
            }

            if (!styleList.ContainsKey(styleIndex))
                styleList.Add(styleIndex, xlCell.Style);
        }

        /// <summary>
        /// Parses the cell value for normal or rich text
        /// Input element should either be a shared string or inline string
        /// </summary>
        /// <param name="element">The element (either a shared string or inline string)</param>
        /// <param name="xlCell">The cell.</param>
        private void ParseCellValue(RstType element, XLCell xlCell)
        {
            var runs = element.Elements<Run>();
            var phoneticRuns = element.Elements<PhoneticRun>();
            var phoneticProperties = element.Elements<PhoneticProperties>();
            Boolean hasRuns = false;
            foreach (Run run in runs)
            {
                var runProperties = run.RunProperties;
                String text = run.Text.InnerText.FixNewLines();

                if (runProperties == null)
                    xlCell.RichText.AddText(text, xlCell.Style.Font);
                else
                {
                    var rt = xlCell.RichText.AddText(text);
                    LoadFont(runProperties, rt);
                }
                if (!hasRuns)
                    hasRuns = true;
            }

            if (!hasRuns)
                xlCell.SetInternalCellValueString(XmlEncoder.DecodeString(element.Text.InnerText));

            #region Load PhoneticProperties

            var pp = phoneticProperties.FirstOrDefault();
            if (pp != null)
            {
                if (pp.Alignment != null)
                    xlCell.RichText.Phonetics.Alignment = pp.Alignment.Value.ToClosedXml();
                if (pp.Type != null)
                    xlCell.RichText.Phonetics.Type = pp.Type.Value.ToClosedXml();

                LoadFont(pp, xlCell.RichText.Phonetics);
            }

            #endregion Load PhoneticProperties

            #region Load Phonetic Runs

            foreach (PhoneticRun pr in phoneticRuns)
            {
                xlCell.RichText.Phonetics.Add(pr.Text.InnerText.FixNewLines(), (Int32)pr.BaseTextStartIndex.Value,
                                              (Int32)pr.EndingBaseIndex.Value);
            }

            #endregion Load Phonetic Runs
        }

        private void LoadNumberFormat(NumberingFormat nfSource, IXLNumberFormat nf)
        {
            if (nfSource == null) return;

            if (nfSource.NumberFormatId != null && nfSource.NumberFormatId.Value < XLConstants.NumberOfBuiltInStyles)
                nf.NumberFormatId = (Int32)nfSource.NumberFormatId.Value;
            else if (nfSource.FormatCode != null)
                nf.Format = nfSource.FormatCode.Value;
        }

        private void LoadBorder(Border borderSource, IXLBorder border)
        {
            if (borderSource == null) return;

            LoadBorderValues(borderSource.DiagonalBorder, border.SetDiagonalBorder, border.SetDiagonalBorderColor);

            if (borderSource.DiagonalUp != null)
                border.DiagonalUp = borderSource.DiagonalUp.Value;
            if (borderSource.DiagonalDown != null)
                border.DiagonalDown = borderSource.DiagonalDown.Value;

            LoadBorderValues(borderSource.LeftBorder, border.SetLeftBorder, border.SetLeftBorderColor);
            LoadBorderValues(borderSource.RightBorder, border.SetRightBorder, border.SetRightBorderColor);
            LoadBorderValues(borderSource.TopBorder, border.SetTopBorder, border.SetTopBorderColor);
            LoadBorderValues(borderSource.BottomBorder, border.SetBottomBorder, border.SetBottomBorderColor);
        }

        private void LoadBorderValues(BorderPropertiesType source, Func<XLBorderStyleValues, IXLStyle> setBorder, Func<XLColor, IXLStyle> setColor)
        {
            if (source != null)
            {
                if (source.Style != null)
                    setBorder(source.Style.Value.ToClosedXml());
                if (source.Color != null)
                    setColor(GetColor(source.Color));
            }
        }

        // Differential fills store the patterns differently than other fills
        // Actually differential fills make more sense. bg is bg and fg is fg
        // 'Other' fills store the bg color in the fg field when pattern type is solid
        private void LoadFill(Fill openXMLFill, IXLFill closedXMLFill, Boolean differentialFillFormat)
        {
            if (openXMLFill == null || openXMLFill.PatternFill == null) return;

            if (openXMLFill.PatternFill.PatternType != null)
                closedXMLFill.PatternType = openXMLFill.PatternFill.PatternType.Value.ToClosedXml();
            else
                closedXMLFill.PatternType = XLFillPatternValues.Solid;

            switch (closedXMLFill.PatternType)
            {
                case XLFillPatternValues.None:
                    break;

                case XLFillPatternValues.Solid:
                    if (differentialFillFormat)
                    {
                        if (openXMLFill.PatternFill.BackgroundColor != null)
                            closedXMLFill.BackgroundColor = GetColor(openXMLFill.PatternFill.BackgroundColor);
                    }
                    else
                    {
                        // yes, source is foreground!
                        if (openXMLFill.PatternFill.ForegroundColor != null)
                            closedXMLFill.BackgroundColor = GetColor(openXMLFill.PatternFill.ForegroundColor);
                    }
                    break;

                default:
                    if (openXMLFill.PatternFill.ForegroundColor != null)
                        closedXMLFill.PatternColor = GetColor(openXMLFill.PatternFill.ForegroundColor);

                    if (openXMLFill.PatternFill.BackgroundColor != null)
                        closedXMLFill.BackgroundColor = GetColor(openXMLFill.PatternFill.BackgroundColor);
                    break;
            }
        }

        private void LoadFont(OpenXmlElement fontSource, IXLFontBase fontBase)
        {
            if (fontSource == null) return;

            fontBase.Bold = GetBoolean(fontSource.Elements<Bold>().FirstOrDefault());
            var fontColor = GetColor(fontSource.Elements<DocumentFormat.OpenXml.Spreadsheet.Color>().FirstOrDefault());
            if (fontColor.HasValue)
                fontBase.FontColor = fontColor;

            var fontFamilyNumbering =
                fontSource.Elements<DocumentFormat.OpenXml.Spreadsheet.FontFamily>().FirstOrDefault();
            if (fontFamilyNumbering != null && fontFamilyNumbering.Val != null)
                fontBase.FontFamilyNumbering =
                    (XLFontFamilyNumberingValues)Int32.Parse(fontFamilyNumbering.Val.ToString());
            var runFont = fontSource.Elements<RunFont>().FirstOrDefault();
            if (runFont != null)
            {
                if (runFont.Val != null)
                    fontBase.FontName = runFont.Val;
            }
            var fontSize = fontSource.Elements<FontSize>().FirstOrDefault();
            if (fontSize != null)
            {
                if ((fontSize).Val != null)
                    fontBase.FontSize = (fontSize).Val;
            }

            fontBase.Italic = GetBoolean(fontSource.Elements<Italic>().FirstOrDefault());
            fontBase.Shadow = GetBoolean(fontSource.Elements<Shadow>().FirstOrDefault());
            fontBase.Strikethrough = GetBoolean(fontSource.Elements<Strike>().FirstOrDefault());

            var underline = fontSource.Elements<Underline>().FirstOrDefault();
            if (underline != null)
            {
                fontBase.Underline = underline.Val != null ? underline.Val.Value.ToClosedXml() : XLFontUnderlineValues.Single;
            }

            var verticalTextAlignment = fontSource.Elements<VerticalTextAlignment>().FirstOrDefault();

            if (verticalTextAlignment == null) return;

            fontBase.VerticalAlignment = verticalTextAlignment.Val != null ? verticalTextAlignment.Val.Value.ToClosedXml() : XLFontVerticalTextAlignmentValues.Baseline;
        }

        private Int32 lastRow;

        private void LoadRows(Stylesheet s, NumberingFormats numberingFormats, Fills fills, Borders borders, Fonts fonts,
                              XLWorksheet ws, SharedStringItem[] sharedStrings,
                              Dictionary<uint, string> sharedFormulasR1C1, Dictionary<Int32, IXLStyle> styleList,
                              Row row)
        {
            Int32 rowIndex = row.RowIndex == null ? ++lastRow : (Int32)row.RowIndex.Value;
            var xlRow = ws.Row(rowIndex, false);

            if (row.Height != null)
                xlRow.Height = row.Height;
            else
            {
                xlRow.Loading = true;
                xlRow.Height = ws.RowHeight;
                xlRow.Loading = false;
            }

            if (row.Hidden != null && row.Hidden)
                xlRow.Hide();

            if (row.Collapsed != null && row.Collapsed)
                xlRow.Collapsed = true;

            if (row.OutlineLevel != null && row.OutlineLevel > 0)
                xlRow.OutlineLevel = row.OutlineLevel;

            if (row.CustomFormat != null)
            {
                Int32 styleIndex = row.StyleIndex != null ? Int32.Parse(row.StyleIndex.InnerText) : -1;
                if (styleIndex > 0)
                {
                    ApplyStyle(xlRow, styleIndex, s, fills, borders, fonts, numberingFormats);
                }
                else
                {
                    xlRow.Style = DefaultStyle;
                }
            }

            lastCell = 0;
            foreach (Cell cell in row.Elements<Cell>())
                LoadCells(sharedStrings, s, numberingFormats, fills, borders, fonts, sharedFormulasR1C1, ws, styleList,
                          cell, rowIndex);
        }

        private void LoadColumns(Stylesheet s, NumberingFormats numberingFormats, Fills fills, Borders borders,
                                 Fonts fonts, XLWorksheet ws, Columns columns)
        {
            if (columns == null) return;

            var wsDefaultColumn =
                columns.Elements<Column>().FirstOrDefault(c => c.Max == XLHelper.MaxColumnNumber);

            if (wsDefaultColumn != null && wsDefaultColumn.Width != null)
                ws.ColumnWidth = wsDefaultColumn.Width - ColumnWidthOffset;

            Int32 styleIndexDefault = wsDefaultColumn != null && wsDefaultColumn.Style != null
                                          ? Int32.Parse(wsDefaultColumn.Style.InnerText)
                                          : -1;
            if (styleIndexDefault >= 0)
                ApplyStyle(ws, styleIndexDefault, s, fills, borders, fonts, numberingFormats);

            foreach (Column col in columns.Elements<Column>())
            {
                //IXLStylized toApply;
                if (col.Max == XLHelper.MaxColumnNumber) continue;

                var xlColumns = (XLColumns)ws.Columns(col.Min, col.Max);
                if (col.Width != null)
                {
                    Double width = col.Width - ColumnWidthOffset;
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
                if (styleIndex > 0)
                {
                    ApplyStyle(xlColumns, styleIndex, s, fills, borders, fonts, numberingFormats);
                }
                else
                {
                    xlColumns.Style = DefaultStyle;
                }
            }
        }

        private static XLDataType GetDataTypeFromCell(IXLNumberFormat numberFormat)
        {
            var numberFormatId = numberFormat.NumberFormatId;
            if (numberFormatId == 46U)
                return XLDataType.TimeSpan;
            else if ((numberFormatId >= 14 && numberFormatId <= 22) ||
                     (numberFormatId >= 45 && numberFormatId <= 47))
                return XLDataType.DateTime;
            else if (numberFormatId == 49)
                return XLDataType.Text;
            else
            {
                if (!String.IsNullOrWhiteSpace(numberFormat.Format))
                {
                    var dataType = GetDataTypeFromFormat(numberFormat.Format);
                    return dataType.HasValue ? dataType.Value : XLDataType.Number;
                }
                else
                    return XLDataType.Number;
            }
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
                else if (c == '0' || c == '#' || c == '?')
                    return XLDataType.Number;
                else if (c == 'y' || c == 'm' || c == 'd' || c == 'h' || c == 's')
                    return XLDataType.DateTime;
            }
            return null;
        }

        private static void LoadAutoFilter(AutoFilter af, XLWorksheet ws)
        {
            if (af != null)
            {
                ws.Range(af.Reference.Value).SetAutoFilter();
                var autoFilter = ws.AutoFilter;
                LoadAutoFilterSort(af, ws, autoFilter);
                LoadAutoFilterColumns(af, autoFilter);
            }
        }

        private static void LoadAutoFilterColumns(AutoFilter af, XLAutoFilter autoFilter)
        {
            foreach (var filterColumn in af.Elements<FilterColumn>())
            {
                Int32 column = (int)filterColumn.ColumnId.Value + 1;
                if (filterColumn.CustomFilters != null)
                {
                    var filterList = new List<XLFilter>();
                    autoFilter.Column(column).FilterType = XLFilterType.Custom;
                    autoFilter.Filters.Add(column, filterList);
                    XLConnector connector = filterColumn.CustomFilters.And != null && filterColumn.CustomFilters.And.Value ? XLConnector.And : XLConnector.Or;

                    Boolean isText = false;
                    foreach (var filter in filterColumn.CustomFilters.OfType<CustomFilter>())
                    {
                        String val = filter.Val.Value;
                        if (!Double.TryParse(val, out Double dTest))
                        {
                            isText = true;
                            break;
                        }
                    }

                    foreach (var filter in filterColumn.CustomFilters.OfType<CustomFilter>())
                    {
                        var xlFilter = new XLFilter { Value = filter.Val.Value, Connector = connector };
                        if (isText)
                            xlFilter.Value = filter.Val.Value;
                        else
                            xlFilter.Value = Double.Parse(filter.Val.Value, CultureInfo.InvariantCulture);

                        if (filter.Operator != null)
                            xlFilter.Operator = filter.Operator.Value.ToClosedXml();
                        else
                            xlFilter.Operator = XLFilterOperator.Equal;

                        Func<Object, Boolean> condition = null;
                        switch (xlFilter.Operator)
                        {
                            case XLFilterOperator.Equal:
                                if (isText)
                                    condition = o => o.ToString().Equals(xlFilter.Value.ToString(), StringComparison.OrdinalIgnoreCase);
                                else
                                    condition = o => (o as IComparable).CompareTo(xlFilter.Value) == 0;
                                break;

                            case XLFilterOperator.EqualOrGreaterThan: condition = o => (o as IComparable).CompareTo(xlFilter.Value) >= 0; break;
                            case XLFilterOperator.EqualOrLessThan: condition = o => (o as IComparable).CompareTo(xlFilter.Value) <= 0; break;
                            case XLFilterOperator.GreaterThan: condition = o => (o as IComparable).CompareTo(xlFilter.Value) > 0; break;
                            case XLFilterOperator.LessThan: condition = o => (o as IComparable).CompareTo(xlFilter.Value) < 0; break;
                            case XLFilterOperator.NotEqual:
                                if (isText)
                                    condition = o => !o.ToString().Equals(xlFilter.Value.ToString(), StringComparison.OrdinalIgnoreCase);
                                else
                                    condition = o => (o as IComparable).CompareTo(xlFilter.Value) != 0;
                                break;
                        }

                        xlFilter.Condition = condition;
                        filterList.Add(xlFilter);
                    }
                }
                else if (filterColumn.Filters != null)
                {
                    if (filterColumn.Filters.Elements().All(element => element is Filter))
                        autoFilter.Column(column).FilterType = XLFilterType.Regular;
                    else if (filterColumn.Filters.Elements().All(element => element is DateGroupItem))
                        autoFilter.Column(column).FilterType = XLFilterType.DateTimeGrouping;
                    else
                        throw new NotSupportedException(String.Format("Mixing regular filters and date group filters in a single autofilter column is not supported. Column {0} of {1}", column, autoFilter.Range.ToString()));

                    var filterList = new List<XLFilter>();

                    autoFilter.Filters.Add((int)filterColumn.ColumnId.Value + 1, filterList);

                    Boolean isText = false;
                    foreach (var filter in filterColumn.Filters.OfType<Filter>())
                    {
                        String val = filter.Val.Value;
                        if (!Double.TryParse(val, out Double dTest))
                        {
                            isText = true;
                            break;
                        }
                    }

                    foreach (var filter in filterColumn.Filters.OfType<Filter>())
                    {
                        var xlFilter = new XLFilter { Connector = XLConnector.Or, Operator = XLFilterOperator.Equal };

                        Func<Object, Boolean> condition;
                        if (isText)
                        {
                            xlFilter.Value = filter.Val.Value;
                            condition = o => o.ToString().Equals(xlFilter.Value.ToString(), StringComparison.OrdinalIgnoreCase);
                        }
                        else
                        {
                            xlFilter.Value = Double.Parse(filter.Val.Value, CultureInfo.InvariantCulture);
                            condition = o => (o as IComparable).CompareTo(xlFilter.Value) == 0;
                        }

                        xlFilter.Condition = condition;
                        filterList.Add(xlFilter);
                    }

                    foreach (var dateGroupItem in filterColumn.Filters.OfType<DateGroupItem>())
                    {
                        bool valid = true;

                        if (!(dateGroupItem.DateTimeGrouping?.HasValue ?? false))
                            continue;

                        var xlDateGroupFilter = new XLFilter
                        {
                            Connector = XLConnector.Or,
                            Operator = XLFilterOperator.Equal,
                            DateTimeGrouping = dateGroupItem.DateTimeGrouping?.Value.ToClosedXml() ?? XLDateTimeGrouping.Year
                        };

                        int year = 1900;
                        int month = 1;
                        int day = 1;
                        int hour = 0;
                        int minute = 0;
                        int second = 0;

                        if (xlDateGroupFilter.DateTimeGrouping >= XLDateTimeGrouping.Year)
                        {
                            if (dateGroupItem?.Year?.HasValue ?? false)
                                year = (int)dateGroupItem.Year?.Value;
                            else
                                valid &= false;
                        }

                        if (xlDateGroupFilter.DateTimeGrouping >= XLDateTimeGrouping.Month)
                        {
                            if (dateGroupItem?.Month?.HasValue ?? false)
                                month = (int)dateGroupItem.Month?.Value;
                            else
                                valid &= false;
                        }

                        if (xlDateGroupFilter.DateTimeGrouping >= XLDateTimeGrouping.Day)
                        {
                            if (dateGroupItem?.Day?.HasValue ?? false)
                                day = (int)dateGroupItem.Day?.Value;
                            else
                                valid &= false;
                        }

                        if (xlDateGroupFilter.DateTimeGrouping >= XLDateTimeGrouping.Hour)
                        {
                            if (dateGroupItem?.Hour?.HasValue ?? false)
                                hour = (int)dateGroupItem.Hour?.Value;
                            else
                                valid &= false;
                        }

                        if (xlDateGroupFilter.DateTimeGrouping >= XLDateTimeGrouping.Minute)
                        {
                            if (dateGroupItem?.Minute?.HasValue ?? false)
                                minute = (int)dateGroupItem.Minute?.Value;
                            else
                                valid &= false;
                        }

                        if (xlDateGroupFilter.DateTimeGrouping >= XLDateTimeGrouping.Second)
                        {
                            if (dateGroupItem?.Second?.HasValue ?? false)
                                second = (int)dateGroupItem.Second?.Value;
                            else
                                valid &= false;
                        }

                        var date = new DateTime(year, month, day, hour, minute, second);
                        xlDateGroupFilter.Value = date;

                        xlDateGroupFilter.Condition = date2 => XLDateTimeGroupFilteredColumn.IsMatch(date, (DateTime)date2, xlDateGroupFilter.DateTimeGrouping);

                        if (valid)
                            filterList.Add(xlDateGroupFilter);
                    }
                }
                else if (filterColumn.Top10 != null)
                {
                    var xlFilterColumn = autoFilter.Column(column);
                    autoFilter.Filters.Add(column, null);
                    xlFilterColumn.FilterType = XLFilterType.TopBottom;
                    if (filterColumn.Top10.Percent != null && filterColumn.Top10.Percent.Value)
                        xlFilterColumn.TopBottomType = XLTopBottomType.Percent;
                    else
                        xlFilterColumn.TopBottomType = XLTopBottomType.Items;

                    if (filterColumn.Top10.Top != null && !filterColumn.Top10.Top.Value)
                        xlFilterColumn.TopBottomPart = XLTopBottomPart.Bottom;
                    else
                        xlFilterColumn.TopBottomPart = XLTopBottomPart.Top;

                    xlFilterColumn.TopBottomValue = (int)filterColumn.Top10.Val.Value;
                }
                else if (filterColumn.DynamicFilter != null)
                {
                    autoFilter.Filters.Add(column, null);
                    var xlFilterColumn = autoFilter.Column(column);
                    xlFilterColumn.FilterType = XLFilterType.Dynamic;
                    if (filterColumn.DynamicFilter.Type != null)
                        xlFilterColumn.DynamicType = filterColumn.DynamicFilter.Type.Value.ToClosedXml();
                    else
                        xlFilterColumn.DynamicType = XLFilterDynamicType.AboveAverage;

                    xlFilterColumn.DynamicValue = filterColumn.DynamicFilter.Val.Value;
                }
            }
        }

        private static void LoadAutoFilterSort(AutoFilter af, XLWorksheet ws, IXLBaseAutoFilter autoFilter)
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

        private static void LoadSheetProtection(SheetProtection sp, XLWorksheet ws)
        {
            if (sp == null) return;

            if (sp.Sheet != null) ws.Protection.Protected = sp.Sheet.Value;
            if (sp.Password != null) ws.Protection.PasswordHash = sp.Password.Value;
            if (sp.FormatCells != null) ws.Protection.FormatCells = !sp.FormatCells.Value;
            if (sp.FormatColumns != null) ws.Protection.FormatColumns = !sp.FormatColumns.Value;
            if (sp.FormatRows != null) ws.Protection.FormatRows = !sp.FormatRows.Value;
            if (sp.InsertColumns != null) ws.Protection.InsertColumns = !sp.InsertColumns.Value;
            if (sp.InsertHyperlinks != null) ws.Protection.InsertHyperlinks = !sp.InsertHyperlinks.Value;
            if (sp.InsertRows != null) ws.Protection.InsertRows = !sp.InsertRows.Value;
            if (sp.DeleteColumns != null) ws.Protection.DeleteColumns = !sp.DeleteColumns.Value;
            if (sp.DeleteRows != null) ws.Protection.DeleteRows = !sp.DeleteRows.Value;
            if (sp.AutoFilter != null) ws.Protection.AutoFilter = !sp.AutoFilter.Value;
            if (sp.PivotTables != null) ws.Protection.PivotTables = !sp.PivotTables.Value;
            if (sp.Sort != null) ws.Protection.Sort = !sp.Sort.Value;
            if (sp.Objects != null) ws.Protection.Objects = !sp.Objects.Value;
            if (sp.SelectLockedCells != null) ws.Protection.SelectLockedCells = sp.SelectLockedCells.Value;
            if (sp.SelectUnlockedCells != null) ws.Protection.SelectUnlockedCells = sp.SelectUnlockedCells.Value;
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
                    var dvt = ws.Range(rangeAddress).SetDataValidation();
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

        /// <summary>
        /// Loads the conditional formatting.
        /// </summary>
        // https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.conditionalformattingrule%28v=office.15%29.aspx?f=255&MSPPError=-2147217396
        private void LoadConditionalFormatting(ConditionalFormatting conditionalFormatting, XLWorksheet ws,
            Dictionary<Int32, DifferentialFormat> differentialFormats)
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
                    LoadFont(differentialFormats[(Int32)fr.FormatId.Value].Font, conditionalFormat.Style.Font);
                    LoadFill(differentialFormats[(Int32)fr.FormatId.Value].Fill, conditionalFormat.Style.Fill,
                        differentialFillFormat: true);
                    LoadBorder(differentialFormats[(Int32)fr.FormatId.Value].Border, conditionalFormat.Style.Border);
                    LoadNumberFormat(differentialFormats[(Int32)fr.FormatId.Value].NumberingFormat,
                        conditionalFormat.Style.NumberFormat);
                }

                // The conditional formatting type is compulsory. If it doesn't exist, skip the entire rule.
                if (fr.Type == null) continue;
                conditionalFormat.ConditionalFormatType = fr.Type.Value.ToClosedXml();
                conditionalFormat.OriginalPriority = fr.Priority?.Value ?? Int32.MaxValue;

                if (conditionalFormat.ConditionalFormatType == XLConditionalFormatType.CellIs && fr.Operator != null)
                    conditionalFormat.Operator = fr.Operator.Value.ToClosedXml();

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

                    var id = fr.Descendants<DocumentFormat.OpenXml.Office2010.Excel.Id>().FirstOrDefault();
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
                else
                {
                    foreach (var formula in fr.Elements<Formula>())
                    {
                        if (formula.Text != null
                            && (conditionalFormat.ConditionalFormatType == XLConditionalFormatType.CellIs
                                || conditionalFormat.ConditionalFormatType == XLConditionalFormatType.Expression))
                        {
                            conditionalFormat.Values.Add(GetFormula(formula.Text));
                        }
                    }
                }

                ws.ConditionalFormats.Add(conditionalFormat);
            }
        }

        private void LoadExtensions(WorksheetExtensionList extensions, XLWorksheet ws)
        {
            if (extensions == null)
            {
                return;
            }

            foreach (var conditionalFormattingRule in extensions
                .Descendants<DocumentFormat.OpenXml.Office2010.Excel.ConditionalFormattingRule>()
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
                    var negativeFillColor = conditionalFormattingRule.Descendants<DocumentFormat.OpenXml.Office2010.Excel.NegativeFillColor>().SingleOrDefault();
                    var color = new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = negativeFillColor.Rgb };
                    xlConditionalFormat.Colors.Add(this.GetColor(color));
                }
            }
        }

        private static XLFormula GetFormula(String value)
        {
            var formula = new XLFormula();
            formula._value = value;
            formula.IsFormula = !(value[0] == '"' && value.EndsWith("\""));
            return formula;
        }

        private void ExtractConditionalFormatValueObjects(XLConditionalFormat conditionalFormat, OpenXmlElement element)
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
                conditionalFormat.Colors.Add(GetColor(c));
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
                    xlCell.SettingHyperlink = true;

                    if (hl.Id != null)
                        xlCell.Hyperlink = new XLHyperlink(hyperlinkDictionary[hl.Id], tooltip);
                    else if (hl.Location != null)
                        xlCell.Hyperlink = new XLHyperlink(hl.Location.Value, tooltip);
                    else
                        xlCell.Hyperlink = new XLHyperlink(hl.Reference.Value, tooltip);

                    xlCell.SettingHyperlink = false;
                }
            }
        }

        private static void LoadColumnBreaks(ColumnBreaks columnBreaks, XLWorksheet ws)
        {
            if (columnBreaks == null) return;

            foreach (Break columnBreak in columnBreaks.Elements<Break>().Where(columnBreak => columnBreak.Id != null))
            {
                ws.PageSetup.ColumnBreaks.Add(Int32.Parse(columnBreak.Id.InnerText));
            }
        }

        private static void LoadRowBreaks(RowBreaks rowBreaks, XLWorksheet ws)
        {
            if (rowBreaks == null) return;

            foreach (Break rowBreak in rowBreaks.Elements<Break>())
                ws.PageSetup.RowBreaks.Add(Int32.Parse(rowBreak.Id.InnerText));
        }

        private void LoadSheetProperties(SheetProperties sheetProperty, XLWorksheet ws, out PageSetupProperties pageSetupProperties)
        {
            pageSetupProperties = null;
            if (sheetProperty == null) return;

            if (sheetProperty.TabColor != null)
                ws.TabColor = GetColor(sheetProperty.TabColor);

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
            if (pageSetup.FirstPageNumber != null)
                ws.PageSetup.FirstPageNumber = UInt32.Parse(pageSetup.FirstPageNumber.InnerText);
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
            if (pane == null) return;

            if (pane.State == null ||
                (pane.State != PaneStateValues.FrozenSplit && pane.State != PaneStateValues.Frozen)) return;

            if (pane.HorizontalSplit != null)
                ws.SheetView.SplitColumn = (Int32)pane.HorizontalSplit.Value;
            if (pane.VerticalSplit != null)
                ws.SheetView.SplitRow = (Int32)pane.VerticalSplit.Value;
        }

        private void SetProperties(SpreadsheetDocument dSpreadsheet)
        {
            var p = dSpreadsheet.PackageProperties;
            Properties.Author = p.Creator;
            Properties.Category = p.Category;
            Properties.Comments = p.Description;
            if (p.Created != null)
                Properties.Created = p.Created.Value;
            Properties.Keywords = p.Keywords;
            Properties.LastModifiedBy = p.LastModifiedBy;
            Properties.Status = p.ContentStatus;
            Properties.Subject = p.Subject;
            Properties.Title = p.Title;
        }

        private XLColor GetColor(ColorType color)
        {
            XLColor retVal = null;
            if (color != null)
            {
                if (color.Rgb != null)
                {
                    String htmlColor = "#" + color.Rgb.Value;
                    Color thisColor;
                    if (!_colorList.ContainsKey(htmlColor))
                    {
                        thisColor = ColorStringParser.ParseFromHtml(htmlColor);
                        _colorList.Add(htmlColor, thisColor);
                    }
                    else
                        thisColor = _colorList[htmlColor];
                    retVal = XLColor.FromColor(thisColor);
                }
                else if (color.Indexed != null && color.Indexed <= 64)
                    retVal = XLColor.FromIndex((Int32)color.Indexed.Value);
                else if (color.Theme != null)
                {
                    retVal = color.Tint != null ? XLColor.FromTheme((XLThemeColor)color.Theme.Value, color.Tint.Value) : XLColor.FromTheme((XLThemeColor)color.Theme.Value);
                }
            }
            return retVal ?? XLColor.NoColor;
        }

        private void ApplyStyle(IXLStylized xlStylized, Int32 styleIndex, Stylesheet s, Fills fills, Borders borders,
                                Fonts fonts, NumberingFormats numberingFormats)
        {
            if (s == null) return; //No Stylesheet, no Styles

            var cellFormat = (CellFormat)s.CellFormats.ElementAt(styleIndex);

            var xlStyle = XLStyle.Default.Key;

            if (cellFormat.ApplyProtection != null)
            {
                var protection = cellFormat.Protection;

                if (protection == null)
                    xlStyle.Protection = XLProtectionValue.Default.Key;
                else
                {
                    xlStyle.Protection = new XLProtectionKey
                    {
                        Hidden = protection.Hidden != null && protection.Hidden.HasValue &&
                                                              protection.Hidden.Value,
                        Locked = protection.Locked == null ||
                                (protection.Locked.HasValue && protection.Locked.Value)
                    };
                }
            }

            if (UInt32HasValue(cellFormat.FillId))
            {
                var fill = (Fill)fills.ElementAt((Int32)cellFormat.FillId.Value);
                if (fill.PatternFill != null)
                {
                    LoadFill(fill, xlStylized.InnerStyle.Fill, differentialFillFormat: false);
                }
                xlStyle.Fill = (xlStylized.InnerStyle as XLStyle).Value.Key.Fill;
            }

            var alignment = cellFormat.Alignment;
            if (alignment != null)
            {
                var xlAlignment = xlStyle.Alignment;
                if (alignment.Horizontal != null)
                    xlAlignment.Horizontal = alignment.Horizontal.Value.ToClosedXml();
                if (alignment.Indent != null && alignment.Indent != 0)
                    xlAlignment.Indent = Int32.Parse(alignment.Indent.ToString());
                if (alignment.JustifyLastLine != null)
                    xlAlignment.JustifyLastLine = alignment.JustifyLastLine;
                if (alignment.ReadingOrder != null)
                {
                    xlAlignment.ReadingOrder =
                        (XLAlignmentReadingOrderValues)Int32.Parse(alignment.ReadingOrder.ToString());
                }
                if (alignment.RelativeIndent != null)
                    xlAlignment.RelativeIndent = alignment.RelativeIndent;
                if (alignment.ShrinkToFit != null)
                    xlAlignment.ShrinkToFit = alignment.ShrinkToFit;
                if (alignment.TextRotation != null)
                    xlAlignment.TextRotation = (Int32)alignment.TextRotation.Value;
                if (alignment.Vertical != null)
                    xlAlignment.Vertical = alignment.Vertical.Value.ToClosedXml();
                if (alignment.WrapText != null)
                    xlAlignment.WrapText = alignment.WrapText;

                xlStyle.Alignment = xlAlignment;
            }

            if (UInt32HasValue(cellFormat.BorderId))
            {
                uint borderId = cellFormat.BorderId.Value;
                var border = (Border)borders.ElementAt((Int32)borderId);
                var xlBorder = xlStyle.Border;
                if (border != null)
                {
                    var bottomBorder = border.BottomBorder;
                    if (bottomBorder != null)
                    {
                        if (bottomBorder.Style != null)
                            xlBorder.BottomBorder = bottomBorder.Style.Value.ToClosedXml();

                        var bottomBorderColor = GetColor(bottomBorder.Color);
                        if (bottomBorderColor.HasValue)
                            xlBorder.BottomBorderColor = bottomBorderColor.Key;
                    }
                    var topBorder = border.TopBorder;
                    if (topBorder != null)
                    {
                        if (topBorder.Style != null)
                            xlBorder.TopBorder = topBorder.Style.Value.ToClosedXml();
                        var topBorderColor = GetColor(topBorder.Color);
                        if (topBorderColor.HasValue)
                            xlBorder.TopBorderColor = topBorderColor.Key;
                    }
                    var leftBorder = border.LeftBorder;
                    if (leftBorder != null)
                    {
                        if (leftBorder.Style != null)
                            xlBorder.LeftBorder = leftBorder.Style.Value.ToClosedXml();
                        var leftBorderColor = GetColor(leftBorder.Color);
                        if (leftBorderColor.HasValue)
                            xlBorder.LeftBorderColor = leftBorderColor.Key;
                    }
                    var rightBorder = border.RightBorder;
                    if (rightBorder != null)
                    {
                        if (rightBorder.Style != null)
                            xlBorder.RightBorder = rightBorder.Style.Value.ToClosedXml();
                        var rightBorderColor = GetColor(rightBorder.Color);
                        if (rightBorderColor.HasValue)
                            xlBorder.RightBorderColor = rightBorderColor.Key;
                    }
                    var diagonalBorder = border.DiagonalBorder;
                    if (diagonalBorder != null)
                    {
                        if (diagonalBorder.Style != null)
                            xlBorder.DiagonalBorder = diagonalBorder.Style.Value.ToClosedXml();
                        var diagonalBorderColor = GetColor(diagonalBorder.Color);
                        if (diagonalBorderColor.HasValue)
                            xlBorder.DiagonalBorderColor = diagonalBorderColor.Key;
                        if (border.DiagonalDown != null)
                            xlBorder.DiagonalDown = border.DiagonalDown;
                        if (border.DiagonalUp != null)
                            xlBorder.DiagonalUp = border.DiagonalUp;
                    }

                    xlStyle.Border = xlBorder;
                }
            }

            if (UInt32HasValue(cellFormat.FontId))
            {
                var fontId = cellFormat.FontId;
                var font = (DocumentFormat.OpenXml.Spreadsheet.Font)fonts.ElementAt((Int32)fontId.Value);

                var xlFont = xlStyle.Font;
                if (font != null)
                {
                    xlFont.Bold = GetBoolean(font.Bold);

                    var fontColor = GetColor(font.Color);
                    if (fontColor.HasValue)
                        xlFont.FontColor = fontColor.Key;

                    if (font.FontFamilyNumbering != null && (font.FontFamilyNumbering).Val != null)
                    {
                        xlFont.FontFamilyNumbering =
                            (XLFontFamilyNumberingValues)Int32.Parse((font.FontFamilyNumbering).Val.ToString());
                    }
                    if (font.FontName != null)
                    {
                        if ((font.FontName).Val != null)
                            xlFont.FontName = (font.FontName).Val;
                    }
                    if (font.FontSize != null)
                    {
                        if ((font.FontSize).Val != null)
                            xlFont.FontSize = (font.FontSize).Val;
                    }

                    xlFont.Italic = GetBoolean(font.Italic);
                    xlFont.Shadow = GetBoolean(font.Shadow);
                    xlFont.Strikethrough = GetBoolean(font.Strike);

                    if (font.Underline != null)
                    {
                        xlFont.Underline = font.Underline.Val != null
                                            ? (font.Underline).Val.Value.ToClosedXml()
                                            : XLFontUnderlineValues.Single;
                    }

                    if (font.VerticalTextAlignment != null)
                    {
                        xlFont.VerticalAlignment = font.VerticalTextAlignment.Val != null
                                                    ? (font.VerticalTextAlignment).Val.Value.ToClosedXml()
                                                    : XLFontVerticalTextAlignmentValues.Baseline;
                    }

                    xlStyle.Font = xlFont;
                }
            }

            if (UInt32HasValue(cellFormat.NumberFormatId))
            {
                var numberFormatId = cellFormat.NumberFormatId;

                string formatCode = String.Empty;
                if (numberingFormats != null)
                {
                    var numberingFormat =
                        numberingFormats.FirstOrDefault(
                            nf =>
                            ((NumberingFormat)nf).NumberFormatId != null &&
                            ((NumberingFormat)nf).NumberFormatId.Value == numberFormatId) as NumberingFormat;

                    if (numberingFormat != null && numberingFormat.FormatCode != null)
                        formatCode = numberingFormat.FormatCode.Value;
                }

                var xlNumberFormat = xlStyle.NumberFormat;
                if (formatCode.Length > 0)
                {
                    xlNumberFormat.Format = formatCode;
                    xlNumberFormat.NumberFormatId = -1;
                }
                else
                    xlNumberFormat.NumberFormatId = (Int32)numberFormatId.Value;
                xlStyle.NumberFormat = xlNumberFormat;
            }

            xlStylized.InnerStyle = new XLStyle(xlStylized, xlStyle);
        }

        private static Boolean UInt32HasValue(UInt32Value value)
        {
            return value != null && value.HasValue;
        }

        private static Boolean GetBoolean(BooleanPropertyType property)
        {
            if (property != null)
            {
                if (property.Val != null)
                    return property.Val;
                return true;
            }

            return false;
        }
    }
}
