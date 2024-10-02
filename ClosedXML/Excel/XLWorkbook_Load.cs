#nullable disable

using ClosedXML.Extensions;
using ClosedXML.Utils;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ClosedXML.Excel.IO;
using ClosedXML.Parser;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Formula = DocumentFormat.OpenXml.Spreadsheet.Formula;
using Op = DocumentFormat.OpenXml.CustomProperties;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using static ClosedXML.Excel.XLPredefinedFormat.DateTime;

namespace ClosedXML.Excel
{
    using Ap;
    using DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;
    using Drawings;
    using Op;
    using System.Drawing;

    public partial class XLWorkbook
    {
        private void Load(String file, LoadOptions loadOptions)
        {
            LoadSheets(file, loadOptions);
        }

        private void Load(Stream stream, LoadOptions loadOptions)
        {
            LoadSheets(stream, loadOptions);
        }

        private void LoadSheets(String fileName, LoadOptions loadOptions)
        {
            using (var dSpreadsheet = SpreadsheetDocument.Open(fileName, false))
                LoadSpreadsheetDocument(dSpreadsheet, loadOptions);
        }

        private void LoadSheets(Stream stream, LoadOptions loadOptions)
        {
            using (var dSpreadsheet = SpreadsheetDocument.Open(stream, false))
                LoadSpreadsheetDocument(dSpreadsheet, loadOptions);
        }

        private void LoadSheetsFromTemplate(String fileName, LoadOptions loadOptions)
        {
            using (var dSpreadsheet = SpreadsheetDocument.CreateFromTemplate(fileName))
                LoadSpreadsheetDocument(dSpreadsheet, loadOptions);

            // If we load a workbook as a template, we have to treat it as a "new" workbook.
            // The original file will NOT be copied into place before changes are applied
            // Hence all loaded RelIds have to be cleared
            ResetAllRelIds();
        }

        private void ResetAllRelIds()
        {
            foreach (var pc in PivotCachesInternal)
                pc.WorkbookCacheRelId = null;

            var sheetId = 1u;
            foreach (var ws in WorksheetsInternal)
            {
                // Ensure unique sheetId for each sheet. 
                ws.SheetId = sheetId++;
                ws.RelId = null;

                foreach (var pt in ws.PivotTables.Cast<XLPivotTable>())
                {
                    pt.CacheDefinitionRelId = null;
                    pt.RelId = null;
                }

                foreach (var picture in ws.Pictures.Cast<XLPicture>())
                    picture.RelId = null;

                foreach (var table in ws.Tables.Cast<XLTable>())
                    table.RelId = null;
            }
        }

        private void LoadSpreadsheetDocument(SpreadsheetDocument dSpreadsheet, LoadOptions loadOptions)
        {
            var context = new LoadContext();
            ShapeIdManager = new XLIdManager();
            SetProperties(dSpreadsheet);

            SharedStringItem[] sharedStrings = null;
            var workbookPart = dSpreadsheet.WorkbookPart;
            if (workbookPart.GetPartsOfType<SharedStringTablePart>().Any())
            {
                var shareStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                sharedStrings = shareStringPart.SharedStringTable.Elements<SharedStringItem>().ToArray();
            }

            LoadWorkbookTheme(workbookPart?.ThemePart, this);

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

            var wbProps = workbookPart.Workbook.WorkbookProperties;
            if (wbProps != null)
                Use1904DateSystem = OpenXmlHelper.GetBooleanValueAsBool(wbProps.Date1904, false);

            var wbFilesharing = workbookPart.Workbook.FileSharing;
            if (wbFilesharing != null)
            {
                FileSharing.ReadOnlyRecommended = OpenXmlHelper.GetBooleanValueAsBool(wbFilesharing.ReadOnlyRecommended, false);
                FileSharing.UserName = wbFilesharing.UserName?.Value;
            }

            LoadWorkbookProtection(workbookPart.Workbook.WorkbookProtection, this);

            var calculationProperties = workbookPart.Workbook.CalculationProperties;
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

            Stylesheet s = workbookPart.WorkbookStylesPart?.Stylesheet;
            NumberingFormats numberingFormats = s?.NumberingFormats;
            context.LoadNumberFormats(numberingFormats);
            Fills fills = s?.Fills;
            Borders borders = s?.Borders;
            Fonts fonts = s?.Fonts;
            Int32 dfCount = 0;
            Dictionary<Int32, DifferentialFormat> differentialFormats;
            if (s != null && s.DifferentialFormats != null)
                differentialFormats = s.DifferentialFormats.Elements<DifferentialFormat>().ToDictionary(k => dfCount++);
            else
                differentialFormats = new Dictionary<Int32, DifferentialFormat>();

            // If the loaded workbook has a changed "Normal" style, it might affect the default width of a column.
            var normalStyle = s?.CellStyles?.Elements<CellStyle>().FirstOrDefault(x => x.BuiltinId is not null && x.BuiltinId.Value == 0);
            if (normalStyle != null)
            {
                var normalStyleKey = ((XLStyle)Style).Key;
                LoadStyle(ref normalStyleKey, (Int32)normalStyle.FormatId.Value, s, fills, borders, fonts, numberingFormats);
                Style = new XLStyle(null, normalStyleKey);
                ColumnWidth = CalculateColumnWidth(8, Style.Font, this);
            }

            // We loop through the sheets in 2 passes: first just to add the sheets and second to add all the data for the sheets.
            // We do this mainly because it skips a very costly calculation invalidation step, but it also make things more consistent,
            // e.g. when reading calculations that reference other sheets, we know that those sheets always already exist.
            // That consistency point isn't required yet but could be taken advantage of in the future.
            var sheets = workbookPart.Workbook.Sheets;
            Int32 position = 0;
            foreach (var dSheet in sheets.OfType<Sheet>())
            {
                position++;
                var sheetName = dSheet.Name;
                var sheetId = dSheet.SheetId.Value;

                if (string.IsNullOrEmpty(dSheet.Id))
                {
                    // Some non-Excel producers create sheets with empty relId.
                    var emptySheet = WorksheetsInternal.Add(sheetName, position, sheetId);
                    if (dSheet.State != null)
                        emptySheet.Visibility = dSheet.State.Value.ToClosedXml();

                    continue;
                }

                // Although relationship to worksheet is most common, there can be other types
                // than worksheet, e.g. chartSheet. Since we can't load them, add them to list
                // of unsupported sheets and copy them when saving. See Codeplex #6932.
                var worksheetPart = workbookPart.GetPartById(dSheet.Id) as WorksheetPart;
                if (worksheetPart == null)
                {
                    UnsupportedSheets.Add(new UnsupportedSheet { SheetId = sheetId, Position = position });
                    continue;
                }

                var ws = WorksheetsInternal.Add(sheetName, position, sheetId);
                ws.RelId = dSheet.Id;

                if (dSheet.State != null)
                    ws.Visibility = dSheet.State.Value.ToClosedXml();
            }

            position = 0;
            foreach (var dSheet in sheets.OfType<Sheet>())
            {
                position++;
                var sheetName = dSheet.Name;
                var sheetId = dSheet.SheetId.Value;

                if (string.IsNullOrEmpty(dSheet.Id))
                {
                    // Some non-Excel producers create sheets with empty relId.
                    continue;
                }

                // Although relationship to worksheet is most common, there can be other types
                // than worksheet, e.g. chartSheet. Since we can't load them, add them to list
                // of unsupported sheets and copy them when saving. See Codeplex #6932.
                var worksheetPart = workbookPart.GetPartById(dSheet.Id) as WorksheetPart;
                if (worksheetPart == null)
                {
                    continue;
                }

                var sharedFormulasR1C1 = new Dictionary<UInt32, String>();
                if (!WorksheetsInternal.TryGetWorksheet(sheetName, out var ws))
                {
                    // This shouldn't be possible, as all worksheets should have already been added in the loop before this loop
                    continue;
                }

                ApplyStyle(ws, 0, s, fills, borders, fonts, numberingFormats);

                var styleList = new Dictionary<int, IXLStyle>();// {{0, ws.Style}};
                PageSetupProperties pageSetupProperties = null;

                lastRow = 0;

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
                                    ws.ColumnWidth = XLHelper.ConvertWidthToNoC(sheetFormatProperties.DefaultColumnWidth.Value, ws.Style.Font, this);
                                else if (sheetFormatProperties.BaseColumnWidth != null)
                                    ws.ColumnWidth = CalculateColumnWidth(sheetFormatProperties.BaseColumnWidth.Value, ws.Style.Font, this);
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
                            LoadRow(s, numberingFormats, fills, borders, fonts, ws, sharedStrings, sharedFormulasR1C1,
                                     styleList, reader);
                        }
                        else if (reader.ElementType == typeof(AutoFilter))
                            LoadAutoFilter((AutoFilter)reader.LoadCurrentElement(), ws);
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

                ws.ConditionalFormats.ReorderAccordingToOriginalPriority();

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

                // TODO: add new variant to ThreadedCommentLoading to be able to support threaded comments natively.
                //    Check which elements can be contained in a threaded coment and how comments relate to each other (threads/tree).
                if (loadOptions.ThreadedCommentLoading == ThreadedCommentLoading.ConvertToNotes)
                {
                    if (worksheetPart.WorksheetThreadedCommentsParts != null)
                    {
                        foreach (var threadedCommentPart in worksheetPart.WorksheetThreadedCommentsParts)
                        {
                            foreach (var threadedComment in threadedCommentPart.RootElement.Elements<ThreadedComment> ())
                            {
                                // find cell by reference
                                var cell = ws.Cell (threadedComment.Ref);

                                // author etc. is handled by the backward-copmatible entries WorksheetCommentsPart
                                foreach (var threadedCommentText in threadedComment.Elements<ThreadedCommentText> ())
                                {
                                    var xlComment = cell.GetComment() ?? cell.CreateComment ();

                                    xlComment.AddText (threadedCommentText.InnerText.FixNewLines ());
                                }
                            }
                        }
                    }
                }

                if (worksheetPart.WorksheetCommentsPart != null)
                {
                    var root = worksheetPart.WorksheetCommentsPart.Comments;
                    var authors = root.GetFirstChild<Authors>().ChildElements.OfType<Author>().ToList();
                    var comments = root.GetFirstChild<CommentList>().ChildElements.OfType<Comment>().ToList();

                    // **** MAYBE FUTURE SHAPE SIZE SUPPORT
                    var shapes = GetCommentShapes(worksheetPart);

                    for (var i = 0; i < comments.Count; i++)
                    {
                        var c = comments[i];

                        XElement shape = null;
                        if (i < shapes.Count)
                            shape = shapes[i];

                        // find cell by reference
                        var cell = ws.Cell(c.Reference);

                        var shapeIdString = shape?.Attribute("id")?.Value;
                        if (shapeIdString?.StartsWith("_x0000_s") ?? false)
                            shapeIdString = shapeIdString.Substring(8);

                        int? shapeId = int.TryParse(shapeIdString, out int sid) ? (int?)sid : null;

                        XLComment xlComment = null;
                        if (loadOptions.ThreadedCommentLoading == ThreadedCommentLoading.ConvertToNotes)
                        {
                            xlComment = cell.GetComment () ?? cell.CreateComment (shapeId);

                             if (shapeId != null)
                                xlComment.ShapeId = shapeId.Value;
                        }
                        else
                        {
                            xlComment = cell.CreateComment(shapeId);
                        }

                        xlComment.Author = authors[(int)c.AuthorId.Value].InnerText;
                        ShapeIdManager.Add(xlComment.ShapeId);

                        var commentText = c.GetFirstChild<CommentText> ();
                        var runs = commentText.Elements<Run> ();
                        foreach (var run in runs)
                        {
                            var runProperties = run.RunProperties;
                            String text = run.Text.InnerText.FixNewLines();
                            var rt = xlComment.AddText(text);
                            OpenXmlHelper.LoadFont(runProperties, rt);
                        }

                        if (shape != null)
                        {
                            LoadShapeProperties(xlComment, shape);

                            var clientData = shape.Elements().First(e => e.Name.LocalName == "ClientData");
                            LoadClientData(xlComment, clientData);

                            var textBox = shape.Elements().First(e => e.Name.LocalName == "textbox");
                            LoadTextBox(xlComment, textBox);

                            var alt = shape.Attribute("alt");
                            if (alt != null) xlComment.Style.Web.SetAlternateText(alt.Value);

                            LoadColorsAndLines(xlComment, shape);

                            //var insetmode = (string)shape.Attributes().First(a=> a.Name.LocalName == "insetmode");
                            //xlComment.Style.Margins.Automatic = insetmode != null && insetmode.Equals("auto");
                        }
                    }
                }

                #endregion LoadComments
            }

            var workbook = workbookPart.Workbook;

            var bookViews = workbook.BookViews;
            if (bookViews != null && bookViews.FirstOrDefault() is WorkbookView workbookView)
            {
                if (workbookView.ActiveTab == null || !workbookView.ActiveTab.HasValue)
                {
                    Worksheets.First().SetTabActive().Unhide();
                }
                else
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

            PivotTableCacheDefinitionPartReader.Load(workbookPart, this);

            // Delay loading of pivot tables until all sheets have been loaded
            foreach (var dSheet in sheets.OfType<Sheet>())
            {
                if (string.IsNullOrEmpty(dSheet.Id))
                {
                    // Some non-Excel producers create sheets with empty relId.
                    continue;
                }

                // The referenced sheet can also be ChartsheetPart. Only look for pivot tables in normal sheet parts.
                var worksheetPart = workbookPart.GetPartById(dSheet.Id) as WorksheetPart;

                if (worksheetPart is not null)
                {
                    var ws = (XLWorksheet)WorksheetsInternal.Worksheet(dSheet.Name);

                    foreach (var pivotTablePart in worksheetPart.PivotTableParts)
                    {
                        PivotTableDefinitionPartReader.Load(workbookPart, differentialFormats, pivotTablePart, worksheetPart, ws, context);
                    }
                }
            }
        }

        /// <summary>
        /// Calculate expected column width as a number displayed in the column in Excel from
        /// number of characters that should fit into the width and a font.
        /// </summary>
        internal static double CalculateColumnWidth(double charWidth, IXLFont font, XLWorkbook workbook)
        {
            // Convert width as a number of characters and translate it into a given number of pixels.
            int mdw = workbook.GraphicEngine.GetMaxDigitWidth(font, workbook.DpiX).RoundToInt();
            int defaultColWidthPx = XLHelper.NoCToPixels(charWidth, mdw).RoundToInt();

            // Excel then rounds this number up to the nearest multiple of 8 pixels, so that
            // scrolling across columns and rows is faster.
            int roundUpToMultiple = defaultColWidthPx + (8 - defaultColWidthPx % 8);

            // and last convert the width in pixels to width displayed in Excel. Shouldn't round the number, because
            // it causes inconsistency with conversion to other units, but other places in ClosedXML do = keep for now.
            double defaultColumnWidth = XLHelper.PixelToNoC(roundUpToMultiple, mdw).Round(2);
            return defaultColumnWidth;
        }

        private void LoadDrawings(WorksheetPart wsPart, XLWorksheet ws)
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

                        var picture = ws.AddPicture(ms, vsdp.Name, Convert.ToInt32(vsdp.Id.Value)) as XLPicture;
                        picture.RelId = imgId;

                        Xdr.ShapeProperties spPr = anchor.Descendants<Xdr.ShapeProperties>().First();
                        picture.Placement = XLPicturePlacement.FreeFloating;

                        if (spPr?.Transform2D?.Extents?.Cx.HasValue ?? false)
                            picture.Width = ConvertFromEnglishMetricUnits(spPr.Transform2D.Extents.Cx, ws.Workbook.DpiX);

                        if (spPr?.Transform2D?.Extents?.Cy.HasValue ?? false)
                            picture.Height = ConvertFromEnglishMetricUnits(spPr.Transform2D.Extents.Cy, ws.Workbook.DpiY);

                        if (anchor is Xdr.AbsoluteAnchor)
                        {
                            var absoluteAnchor = anchor as Xdr.AbsoluteAnchor;
                            picture.MoveTo(
                                ConvertFromEnglishMetricUnits(absoluteAnchor.Position.X.Value, ws.Workbook.DpiX),
                                ConvertFromEnglishMetricUnits(absoluteAnchor.Position.Y.Value, ws.Workbook.DpiY)
                            );
                        }
                        else if (anchor is Xdr.OneCellAnchor)
                        {
                            var oneCellAnchor = anchor as Xdr.OneCellAnchor;
                            var from = LoadMarker(ws, oneCellAnchor.FromMarker);
                            picture.MoveTo(from.Cell, from.Offset);
                        }
                        else if (anchor is Xdr.TwoCellAnchor)
                        {
                            var twoCellAnchor = anchor as Xdr.TwoCellAnchor;
                            var from = LoadMarker(ws, twoCellAnchor.FromMarker);
                            var to = LoadMarker(ws, twoCellAnchor.ToMarker);

                            if (twoCellAnchor.EditAs == null || !twoCellAnchor.EditAs.HasValue || twoCellAnchor.EditAs.Value == Xdr.EditAsValues.TwoCell)
                            {
                                picture.MoveTo(from.Cell, from.Offset, to.Cell, to.Offset);
                            }
                            else if (twoCellAnchor.EditAs.Value == Xdr.EditAsValues.Absolute)
                            {
                                var shapeProperties = twoCellAnchor.Descendants<Xdr.ShapeProperties>().FirstOrDefault();
                                if (shapeProperties != null)
                                {
                                    picture.MoveTo(
                                        ConvertFromEnglishMetricUnits(spPr.Transform2D.Offset.X, ws.Workbook.DpiX),
                                        ConvertFromEnglishMetricUnits(spPr.Transform2D.Offset.Y, ws.Workbook.DpiY)
                                    );
                                }
                            }
                            else if (twoCellAnchor.EditAs.Value == Xdr.EditAsValues.OneCell)
                            {
                                picture.MoveTo(from.Cell, from.Offset);
                            }
                        }
                    }
                }
            }
        }

        private static Int32 ConvertFromEnglishMetricUnits(long emu, double resolution)
        {
            return Convert.ToInt32(emu * resolution / 914400);
        }

        private static XLMarker LoadMarker(XLWorksheet ws, Xdr.MarkerType marker)
        {
            var row = Math.Min(XLHelper.MaxRowNumber, Math.Max(1, Convert.ToInt32(marker.RowId.InnerText) + 1));
            var column = Math.Min(XLHelper.MaxColumnNumber, Math.Max(1, Convert.ToInt32(marker.ColumnId.InnerText) + 1));
            return new XLMarker(
                ws.Cell(row, column),
                new Point(
                    ConvertFromEnglishMetricUnits(Convert.ToInt32(marker.ColumnOffset.InnerText), ws.Workbook.DpiX),
                    ConvertFromEnglishMetricUnits(Convert.ToInt32(marker.RowOffset.InnerText), ws.Workbook.DpiY)
                )
            );
        }

        #region Comment Helpers

        private static IList<XElement> GetCommentShapes(WorksheetPart worksheetPart)
        {
            // Cannot get this to return Vml.Shape elements
            foreach (var vmlPart in worksheetPart.VmlDrawingParts)
            {
                using (var stream = vmlPart.GetStream(FileMode.Open))
                {
                    var xdoc = XDocumentExtensions.Load(stream);
                    if (xdoc == null)
                        continue;

                    var root = xdoc.Root.Element("xml") ?? xdoc.Root;

                    if (root == null)
                        continue;

                    var shapes = root
                        .Elements(XName.Get("shape", "urn:schemas-microsoft-com:vml"))
                        .Where(e => new[]
                        {
                            "#" + XLConstants.Comment.ShapeTypeId ,
                            "#" + XLConstants.Comment.AlternateShapeTypeId
                        }.Contains(e.Attribute("type")?.Value))
                        .ToList();

                    if (shapes != null)
                        return shapes;
                }
            }

            throw new ArgumentException("Could not load comments file");
        }

        #endregion Comment Helpers

        private String GetTableColumnName(string name)
        {
            return name.Replace("_x000a_", Environment.NewLine).Replace("_x005f_x000a_", "_x000a_");
        }

        private void LoadColorsAndLines<T>(IXLDrawing<T> drawing, XElement shape)
        {
            var strokeColor = shape.Attribute("strokecolor");
            if (strokeColor != null) drawing.Style.ColorsAndLines.LineColor = XLColor.FromVmlColor(strokeColor.Value);

            var strokeWeight = shape.Attribute("strokeweight");
            if (strokeWeight != null && TryGetPtValue(strokeWeight.Value, out var lineWeight))
                drawing.Style.ColorsAndLines.LineWeight = lineWeight;

            var fillColor = shape.Attribute("fillcolor");
            if (fillColor != null) drawing.Style.ColorsAndLines.FillColor = XLColor.FromVmlColor(fillColor.Value);

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
            drawing.Visible = visible != null &&
                              (string.IsNullOrEmpty(visible.Value) ||
                               visible.Value.StartsWith("t", StringComparison.OrdinalIgnoreCase));

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
            drawing.Position.ColumnOffset = Double.Parse(location[1], CultureInfo.InvariantCulture) / 7.5;
            drawing.Position.Row = int.Parse(location[2]) + 1;
            drawing.Position.RowOffset = Double.Parse(location[3], CultureInfo.InvariantCulture);
        }

        private void LoadShapeProperties<T>(IXLDrawing<T> xlDrawing, XElement shape)
        {
            if (shape.Attribute("style") == null)
                return;

            foreach (var attributePair in shape.Attribute("style").Value.Split(';'))
            {
                var split = attributePair.Split(':');
                if (split.Length != 2) continue;

                var attribute = split[0].Trim().ToLower();
                var value = split[1].Trim();

                switch (attribute)
                {
                    case "visibility": xlDrawing.Visible = string.Equals("visible", value, StringComparison.OrdinalIgnoreCase); break;
                    case "width":
                        if (TryGetPtValue(value, out var ptWidth))
                        {
                            xlDrawing.Style.Size.Width = ptWidth / 7.5;
                        }
                        break;

                    case "height":
                        if (TryGetPtValue(value, out var ptHeight))
                        {
                            xlDrawing.Style.Size.Height = ptHeight;
                        }
                        break;

                    case "z-index":
                        if (Int32.TryParse(value, out var zOrder))
                        {
                            xlDrawing.ZOrder = zOrder;
                        }
                        break;
                }
            }
        }

        private readonly Dictionary<string, double> knownUnits = new Dictionary<string, double>
        {
            {"pt", 1.0},
            {"in", 72.0},
            {"mm", 72.0/25.4}
        };

        private bool TryGetPtValue(string value, out double result)
        {
            var knownUnit = knownUnits.FirstOrDefault(ku => value.Contains(ku.Key));

            if (knownUnit.Key == null)
                return Double.TryParse(value, out result);

            value = value.Replace(knownUnit.Key, String.Empty);

            if (Double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out result))
            {
                result *= knownUnit.Value;
                return true;
            }

            result = 0d;
            return false;
        }

        private void LoadDefinedNames(Workbook workbook)
        {
            if (workbook.DefinedNames == null) return;

            foreach (var definedName in workbook.DefinedNames.OfType<DefinedName>())
            {
                var name = definedName.Name;
                var visible = true;
                if (definedName.Hidden != null) visible = !BooleanValue.ToBoolean(definedName.Hidden);

                var localSheetId = -1;
                if (definedName.LocalSheetId?.HasValue ?? false) localSheetId = Convert.ToInt32(definedName.LocalSheetId.Value);

                if (name == "_xlnm.Print_Area")
                {
                    var fixedNames = validateDefinedNames(definedName.Text.Split(','));
                    foreach (string area in fixedNames)
                    {
                        if (area.Contains("["))
                        {
                            var ws = WorksheetsInternal.FirstOrDefault<XLWorksheet>(w => w.SheetId == (localSheetId + 1));
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

                    var comment = definedName.Comment;
                    if (localSheetId == -1)
                    {
                        if (DefinedNamesInternal.All<XLDefinedName>(nr => nr.Name != name))
                            DefinedNamesInternal.Add(name, text, comment, validateName: false, validateRangeAddress: false).Visible = visible;
                    }
                    else
                    {
                        if (Worksheet(localSheetId + 1).DefinedNames.All(nr => nr.Name != name))
                            ((XLDefinedNames)Worksheet(localSheetId + 1).DefinedNames).Add(name, text, comment, validateName: false, validateRangeAddress: false).Visible = visible;
                    }
                }
            }
        }

        private static Regex definedNameRegex = new Regex(@"\A('?).*\1!.*\z", RegexOptions.Compiled);

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
            sheetArea = sheetArea.Replace("$", "");

            if (sheetArea.Equals("#REF")) return;
            if (IsColReference(sheetArea))
                WorksheetsInternal.Worksheet(sheetName).PageSetup.SetColumnsToRepeatAtLeft(sheetArea);
            if (IsRowReference(sheetArea))
                WorksheetsInternal.Worksheet(sheetName).PageSetup.SetRowsToRepeatAtTop(sheetArea);
        }

        // either $A:$X => true or $1:$99 => false
        private static bool IsColReference(string sheetArea)
        {
            return sheetArea.All(c => c == ':' || char.IsLetter(c));
        }

        private static bool IsRowReference(string sheetArea)
        {
            return sheetArea.All(c => c == ':' || char.IsNumber(c));
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
                sheetName = string.Join("!", sections.Take(sections.Length - 1)).UnescapeSheetName();
                sheetArea = sections[sections.Length - 1];
            }
        }

        private Int32 lastColumnNumber;

        private void LoadCell(SharedStringItem[] sharedStrings, Stylesheet s, NumberingFormats numberingFormats,
                              Fills fills, Borders borders, Fonts fonts, Dictionary<uint, string> sharedFormulasR1C1,
                              XLWorksheet ws, Dictionary<Int32, IXLStyle> styleList, OpenXmlPartReader reader, Int32 rowIndex)
        {
            Debug.Assert(reader.LocalName == "c" && reader.IsStartElement);

            var attributes = reader.Attributes;

            var styleIndex = attributes.GetIntAttribute("s") ?? 0;

            var cellAddress = attributes.GetCellRefAttribute("r") ?? new XLSheetPoint(rowIndex, lastColumnNumber + 1);
            lastColumnNumber = cellAddress.Column;

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
                ApplyStyle(xlCell, styleIndex, s, fills, borders, fonts, numberingFormats);
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
                formula = SetCellFormula(ws, cellAddress, reader, sharedFormulasR1C1);

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

            if (Use1904DateSystem && xlCell.DataType == XLDataType.DateTime)
            {
                // Internally ClosedXML stores cells as standard 1900-based style
                // so if a workbook is in 1904-format, we do that adjustment here and when saving.
                xlCell.SetOnlyValue(xlCell.GetDateTime().AddDays(1462));
            }

            if (!styleList.ContainsKey(styleIndex))
                styleList.Add(styleIndex, xlCell.Style);
        }

        private static XLCellFormula SetCellFormula(XLWorksheet ws, XLSheetPoint cellAddress, OpenXmlPartReader reader, Dictionary<uint, string> sharedFormulasR1C1)
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
                if (!sharedFormulasR1C1.TryGetValue(sharedIndex, out var sharedR1C1Formula))
                {
                    // Spec: The first formula in a group of shared formulas is saved
                    // in the f element. This is considered the 'master' formula cell.
                    formula = XLCellFormula.NormalA1(formulaText);
                    formulaSlice.Set(cellAddress, formula);

                    // The key reason why Excel hates shared formulas is likely relative addressing and the messy situation it creates
                    var formulaR1C1 = FormulaConverter.ToR1C1(formulaText, cellAddress.Row, cellAddress.Column);
                    sharedFormulasR1C1.Add(sharedIndex, formulaR1C1);
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
                var input1 = attributes.GetCellRefAttribute("r1") ?? throw MissingRequiredAttr("r1");
                if (is2D)
                {
                    // Input 2 is only used for 2D tables
                    var input2Deleted = attributes.GetBoolAttribute("del2", false);
                    var input2 = attributes.GetCellRefAttribute("r2") ?? throw MissingRequiredAttr("r2");
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

        private static readonly string[] DateCellFormats =
        {
            "yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fff", // Format accepted by OpenXML SDK
            "yyyy-MM-ddTHH:mm", "yyyy-MM-dd" // Formats accepted by Excel.
        };

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
                xlCell.SetOnlyValue(XmlEncoder.DecodeString(element.Text?.InnerText));

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

        private Int32 lastRow;

        private void LoadRow(Stylesheet s, NumberingFormats numberingFormats, Fills fills, Borders borders, Fonts fonts,
                              XLWorksheet ws, SharedStringItem[] sharedStrings,
                              Dictionary<uint, string> sharedFormulasR1C1, Dictionary<Int32, IXLStyle> styleList,
                              OpenXmlPartReader reader)
        {
            Debug.Assert(reader.LocalName == "row");

            var attributes = reader.Attributes;
            var rowIndexAttr = attributes.GetAttribute("r");
            var rowIndex = string.IsNullOrEmpty(rowIndexAttr) ? ++lastRow : int.Parse(rowIndexAttr);

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
                    ApplyStyle(xlRow, styleIndex.Value, s, fills, borders, fonts, numberingFormats);
                }
                else
                {
                    xlRow.Style = ws.Style;
                }
            }

            lastColumnNumber = 0;

            // Move from the start element of 'row' forward. We can get cell, extList or end of row.
            reader.MoveAhead();

            while (reader.IsStartElement("c"))
            {
                LoadCell(sharedStrings, s, numberingFormats, fills, borders, fonts, sharedFormulasR1C1, ws, styleList,
                    reader, rowIndex);

                // Move from end element of 'cell' either to next cell, extList start or end of row.
                reader.MoveAhead();
            }

            // In theory, row can also contain extList, just skip them.
            while (reader.IsStartElement("extLst"))
                reader.Skip();
        }

        private void LoadColumns(Stylesheet s, NumberingFormats numberingFormats, Fills fills, Borders borders,
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
                ApplyStyle(ws, styleIndexDefault, s, fills, borders, fonts, numberingFormats);

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
                    ApplyStyle(xlColumns, styleIndex, s, fills, borders, fonts, numberingFormats);
                }
                else
                {
                    xlColumns.Style = ws.Style;
                }
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

        /// <summary>
        /// Loads the conditional formatting.
        /// </summary>
        // https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.conditionalformattingrule%28v=office.15%29.aspx?f=255&MSPPError=-2147217396
        private void LoadConditionalFormatting(ConditionalFormatting conditionalFormatting, XLWorksheet ws,
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
                        .Select(f => GetFormula(f.Text))
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

                var isPivotTableFormatting = conditionalFormatting.Pivot?.Value ?? false;
                if (isPivotTableFormatting)
                    context.AddPivotTableCf(ws.Name, conditionalFormat);
                else
                    ws.ConditionalFormats.Add(conditionalFormat);
            }
        }

        private void LoadExtensions(WorksheetExtensionList extensions, XLWorksheet ws)
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
                    xlConditionalFormat.Colors.Add(negativeFillColor.ToClosedXMLColor());
                }
            }

            foreach (var slg in extensions
                .Descendants<X14.SparklineGroups>()
                .SelectMany(sparklineGroups => sparklineGroups.Descendants<X14.SparklineGroup>()))
            {
                var xlSparklineGroup = (ws.SparklineGroups as XLSparklineGroups).Add();

                if (slg.Formula != null)
                    xlSparklineGroup.DateRange = Range(slg.Formula.Text);

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

        private static void LoadWorkbookTheme(ThemePart tp, XLWorkbook wb)
        {
            if (tp is null)
                return;

            var colorScheme = tp.Theme?.ThemeElements?.ColorScheme;
            if (colorScheme is not null)
            {
                var background1 = colorScheme.Light1Color?.RgbColorModelHex?.Val?.Value;
                if (!string.IsNullOrEmpty(background1))
                {
                    wb.Theme.Background1 = XLColor.FromHexRgb(background1);
                }
                var text1 = colorScheme.Dark1Color?.RgbColorModelHex?.Val?.Value;
                if (!string.IsNullOrEmpty(text1))
                {
                    wb.Theme.Text1 = XLColor.FromHexRgb(text1);
                }
                var background2 = colorScheme.Light2Color?.RgbColorModelHex?.Val?.Value;
                if (!string.IsNullOrEmpty(background2))
                {
                    wb.Theme.Background2 = XLColor.FromHexRgb(background2);
                }
                var text2 = colorScheme.Dark2Color?.RgbColorModelHex?.Val?.Value;
                if (!string.IsNullOrEmpty(text2))
                {
                    wb.Theme.Text2 = XLColor.FromHexRgb(text2);
                }
                var accent1 = colorScheme.Accent1Color?.RgbColorModelHex?.Val?.Value;
                if (!string.IsNullOrEmpty(accent1))
                {
                    wb.Theme.Accent1 = XLColor.FromHexRgb(accent1);
                }
                var accent2 = colorScheme.Accent2Color?.RgbColorModelHex?.Val?.Value;
                if (!string.IsNullOrEmpty(accent2))
                {
                    wb.Theme.Accent2 = XLColor.FromHexRgb(accent2);
                }
                var accent3 = colorScheme.Accent3Color?.RgbColorModelHex?.Val?.Value;
                if (!string.IsNullOrEmpty(accent3))
                {
                    wb.Theme.Accent3 = XLColor.FromHexRgb(accent3);
                }
                var accent4 = colorScheme.Accent4Color?.RgbColorModelHex?.Val?.Value;
                if (!string.IsNullOrEmpty(accent4))
                {
                    wb.Theme.Accent4 = XLColor.FromHexRgb(accent4);
                }
                var accent5 = colorScheme.Accent5Color?.RgbColorModelHex?.Val?.Value;
                if (!string.IsNullOrEmpty(accent5))
                {
                    wb.Theme.Accent5 = XLColor.FromHexRgb(accent5);
                }
                var accent6 = colorScheme.Accent6Color?.RgbColorModelHex?.Val?.Value;
                if (!string.IsNullOrEmpty(accent6))
                {
                    wb.Theme.Accent6 = XLColor.FromHexRgb(accent6);
                }
                var hyperlink = colorScheme.Hyperlink?.RgbColorModelHex?.Val?.Value;
                if (!string.IsNullOrEmpty(hyperlink))
                {
                    wb.Theme.Hyperlink = XLColor.FromHexRgb(hyperlink);
                }
                var followedHyperlink = colorScheme.FollowedHyperlinkColor?.RgbColorModelHex?.Val?.Value;
                if (!string.IsNullOrEmpty(followedHyperlink))
                {
                    wb.Theme.FollowedHyperlink = XLColor.FromHexRgb(followedHyperlink);
                }
            }
        }

        private static void LoadWorkbookProtection(WorkbookProtection wp, XLWorkbook wb)
        {
            if (wp == null) return;

            wb.Protection.IsProtected = true;

            var algorithmName = wp.WorkbookAlgorithmName?.Value ?? string.Empty;
            if (String.IsNullOrEmpty(algorithmName))
            {
                wb.Protection.PasswordHash = wp.WorkbookPassword?.Value ?? string.Empty;
                wb.Protection.Base64EncodedSalt = string.Empty;
            }
            else if (DescribedEnumParser<XLProtectionAlgorithm.Algorithm>.IsValidDescription(algorithmName))
            {
                wb.Protection.Algorithm = DescribedEnumParser<XLProtectionAlgorithm.Algorithm>.FromDescription(algorithmName);
                wb.Protection.PasswordHash = wp.WorkbookHashValue?.Value ?? string.Empty;
                wb.Protection.SpinCount = wp.WorkbookSpinCount?.Value ?? 0;
                wb.Protection.Base64EncodedSalt = wp.WorkbookSaltValue?.Value ?? string.Empty;
            }

            wb.Protection.AllowElement(XLWorkbookProtectionElements.Structure, !OpenXmlHelper.GetBooleanValueAsBool(wp.LockStructure, false));
            wb.Protection.AllowElement(XLWorkbookProtectionElements.Windows, !OpenXmlHelper.GetBooleanValueAsBool(wp.LockWindows, false));
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
                conditionalFormat.Colors.Add(c.ToClosedXMLColor());
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
            if (pageSetup.FirstPageNumber?.HasValue ?? false)
                ws.PageSetup.FirstPageNumber = (int)pageSetup.FirstPageNumber.Value;
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

        private void SetProperties(SpreadsheetDocument dSpreadsheet)
        {
            var p = dSpreadsheet.PackageProperties;
            Properties.Author = p.Creator;
            Properties.Category = p.Category;
            Properties.Comments = p.Description;
            if (p.Created != null)
                Properties.Created = p.Created.Value;
            if (p.Modified != null)
                Properties.Modified = p.Modified.Value;
            Properties.Keywords = p.Keywords;
            Properties.LastModifiedBy = p.LastModifiedBy;
            Properties.Status = p.ContentStatus;
            Properties.Subject = p.Subject;
            Properties.Title = p.Title;
        }

        private void ApplyStyle(IXLStylized xlStylized, Int32 styleIndex, Stylesheet s, Fills fills, Borders borders,
            Fonts fonts, NumberingFormats numberingFormats)
        {
            var xlStyleKey = XLStyle.Default.Key;
            LoadStyle(ref xlStyleKey, styleIndex, s, fills, borders, fonts, numberingFormats);

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

        private void LoadStyle(ref XLStyleKey xlStyle, Int32 styleIndex, Stylesheet s, Fills fills, Borders borders,
                                Fonts fonts, NumberingFormats numberingFormats)
        {
            if (s == null || s.CellFormats is null) return; //No Stylesheet, no Styles

            var cellFormat = (CellFormat)s.CellFormats.ElementAt(styleIndex);

            var xlIncludeQuotePrefix = OpenXmlHelper.GetBooleanValueAsBool(cellFormat.QuotePrefix, false);
            xlStyle = xlStyle with { IncludeQuotePrefix = xlIncludeQuotePrefix };

            if (cellFormat.ApplyProtection != null)
            {
                var protection = cellFormat.Protection;
                var xlProtection = XLProtectionValue.Default.Key;
                if (protection is not null)
                    xlProtection = OpenXmlHelper.ProtectionToClosedXml(protection, xlProtection);

                xlStyle = xlStyle with { Protection = xlProtection };
            }

            if (UInt32HasValue(cellFormat.FillId))
            {
                var fill = (Fill)fills.ElementAt((Int32)cellFormat.FillId.Value);
                if (fill.PatternFill != null)
                {
                    var xlFill = new XLFill();
                    OpenXmlHelper.LoadFill(fill, xlFill, differentialFillFormat: false);
                    xlStyle = xlStyle with { Fill = xlFill.Key };
                }
            }

            var alignment = cellFormat.Alignment;
            if (alignment != null)
            {
                var xlAlignment = OpenXmlHelper.AlignmentToClosedXml(alignment, xlStyle.Alignment);
                xlStyle = xlStyle with { Alignment = xlAlignment };
            }

            if (UInt32HasValue(cellFormat.BorderId))
            {
                uint borderId = cellFormat.BorderId.Value;
                var border = (Border)borders.ElementAt((Int32)borderId);
                if (border is not null)
                {
                    var xlBorder = OpenXmlHelper.BorderToClosedXml(border, xlStyle.Border);
                    xlStyle = xlStyle with { Border = xlBorder };
                }
            }

            if (UInt32HasValue(cellFormat.FontId))
            {
                var fontId = cellFormat.FontId;
                var font = (DocumentFormat.OpenXml.Spreadsheet.Font)fonts.ElementAt((Int32)fontId.Value);
                if (font is not null)
                {
                    var xlFont = OpenXmlHelper.FontToClosedXml(font, xlStyle.Font);
                    xlStyle = xlStyle with { Font = xlFont };
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
                    xlNumberFormat = XLNumberFormatKey.ForFormat(formatCode);
                }
                else
                    xlNumberFormat = xlNumberFormat with { NumberFormatId = (Int32)numberFormatId.Value };
                xlStyle = xlStyle with { NumberFormat = xlNumberFormat };
            }
        }

        private static Boolean UInt32HasValue(UInt32Value value)
        {
            return value != null && value.HasValue;
        }

        private static Exception MissingRequiredAttr(string attributeName)
        {
            throw new InvalidOperationException($"XML doesn't contain required attribute '{attributeName}'.");
        }
    }
}
