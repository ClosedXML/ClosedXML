#nullable disable

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
using ClosedXML.Excel.IO;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Op = DocumentFormat.OpenXml.CustomProperties;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace ClosedXML.Excel
{
    using Ap;
    using ClosedXML.IO;
    using Drawings;
    using Op;
    using System.Drawing;

    public partial class XLWorkbook
    {
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

        private void LoadSpreadsheetDocument(SpreadsheetDocument dSpreadsheet)
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

            var stylesPart = workbookPart.WorkbookStylesPart;
            if (stylesPart is not null)
            {
                using var xmlReader = CreateTreeReader(stylesPart);
                var stylesReader = new StylesReader(xmlReader, Styles);
                stylesReader.Load();
            }

            Stylesheet s = stylesPart?.Stylesheet;
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
                LoadStyle(ref normalStyleKey, (Int32)normalStyle.FormatId.Value, Styles);
                Style = new XLStyle(null, normalStyleKey);
                ColumnWidth = XLHelper.CalculateColumnWidth(8, Style.Font, this);
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

                if (!WorksheetsInternal.TryGetWorksheet(sheetName, out var ws))
                {
                    // This shouldn't be possible, as all worksheets should have already been added in the loop before this loop
                    continue;
                }

                var worksheetPartReader = new WorksheetPartReader();
                worksheetPartReader.LoadWorksheet(ws, s, worksheetPart, sharedStrings, differentialFormats, context);

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
                        AutoFilterReader.LoadAutoFilterColumns(dTable.AutoFilter, xlTable.AutoFilter);
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
                        var xlComment = cell.CreateComment(shapeId);

                        xlComment.Author = authors[(int)c.AuthorId.Value].InnerText;
                        ShapeIdManager.Add(xlComment.ShapeId);

                        var runs = c.GetFirstChild<CommentText>().Elements<Run>();
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

            // Read cache definition before table definition
            foreach (var pivotTableCacheDefinitionPart in workbookPart.GetPartsOfType<PivotTableCacheDefinitionPart>())
            {
                var pivotCache = PivotTableCacheDefinitionPartReader.Load(workbookPart, pivotTableCacheDefinitionPart, this);
                if (pivotTableCacheDefinitionPart.PivotTableCacheRecordsPart is { } recordsPart)
                {
                    using var reader = CreateTreeReader(recordsPart);
                    var recordsReader = new PivotCacheRecordsReader(reader, pivotCache);
                    recordsReader.ReadRecordsToCache();
                }
            }

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
            var strokeColor = shape.Attribute(@"strokecolor");
            if (strokeColor is not null)
                drawing.Style.ColorsAndLines.LineColor = XLColor.FromVmlColor(strokeColor.Value);

            var strokeWeight = shape.Attribute(@"strokeweight");
            if (strokeWeight != null && TryGetPtValue(strokeWeight.Value, out var lineWeight))
                drawing.Style.ColorsAndLines.LineWeight = lineWeight;

            var fillColor = shape.Attribute(@"fillcolor");
            if (fillColor is not null)
                drawing.Style.ColorsAndLines.FillColor = XLColor.FromVmlColor(fillColor.Value);

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
            if (attStyle != null) LoadTextBoxStyle(xlDrawing, attStyle);

            var attInset = textBox.Attribute("inset");
            if (attInset != null) LoadTextBoxInset(xlDrawing, attInset);
        }

        private void LoadTextBoxInset<T>(IXLDrawing<T> xlDrawing, XAttribute attInset)
        {
            var split = attInset.Value.Split(',');
            xlDrawing.Style.Margins.Left = GetInsetInInches(split[0], DpiX);
            xlDrawing.Style.Margins.Top = GetInsetInInches(split[1], DpiY);
            xlDrawing.Style.Margins.Right = GetInsetInInches(split[2], DpiX);
            xlDrawing.Style.Margins.Bottom = GetInsetInInches(split[3], DpiY);
        }

        /// <summary>
        /// List of all VML length units and their conversion. Key is a name, value is a conversion
        /// function to EMU. See <a href="https://learn.microsoft.com/en-us/windows/win32/vml/msdn-online-vml-units">documentation</a>.
        /// </summary>
        /// <remarks>
        /// OI-29500 says <em>Office also uses EMUs throughout VML as a valid unit system</em>.
        /// Relative units conversions are guesstimated by how Excel 2022 behaves for inset
        /// attribute of <c>TextBox</c> element of a note/comment. Generally speaking, Excel
        /// converts relative values to physical length (e.g. <c>px</c> to <c>pt</c>) and saves
        /// them as such. The <c>ex</c>/<c>em</c> units are not interpreted as described in the
        /// doc, but as 1/90th or an inch. The <c>%</c> seems to be always 0.
        /// </remarks>
        private static readonly Dictionary<string, Func<double, double, Emu?>> VmlLengthUnits = new()
        {
            {"in", (value, _) => Emu.From(value, AbsLengthUnit.Inch) },
            {"cm", (value, _) => Emu.From(value, AbsLengthUnit.Centimeter) },
            {"mm", (value, _) => Emu.From(value, AbsLengthUnit.Millimeter) },
            {"pt", (value, _) => Emu.From(value, AbsLengthUnit.Point) },
            {"pc", (value, _) => Emu.From(value, AbsLengthUnit.Pica) },
            {"emu", (value, _) => Emu.From(value , AbsLengthUnit.Emu) },
            {"px", (value, dpi) => Emu.From(value / dpi, AbsLengthUnit.Inch) },
            {"em", (value, _) => Emu.From(value * 72.0 / 90.0, AbsLengthUnit.Point) },
            {"ex", (value, _) => Emu.From(value * 72.0 / 90.0, AbsLengthUnit.Point) },
            {"%", (_, _) => Emu.ZeroPt },
        };

        private static double GetInsetInInches(string value, double dpi)
        {
            var unit = value.Trim();
            foreach (var (unitName, conversion) in VmlLengthUnits)
            {
                if (unit.EndsWith(unitName) && Double.TryParse(unit[..^unitName.Length], NumberStyles.Float, CultureInfo.InvariantCulture, out var unitValue))
                {
                    var insetEmu = conversion(unitValue, dpi) ?? Emu.ZeroPt;
                    return insetEmu.To(AbsLengthUnit.Inch);
                }
            }

            // Excel treats no/unexpected unit as 0
            return 0;
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

        internal static void LoadStyle(ref XLStyleKey xlStyle, Int32 styleIndex, XLWorkbookStyles styles)
        {
            var xlCellFormat = styles.CellFormats[styleIndex];
            xlStyle = xlStyle with { IncludeQuotePrefix = xlCellFormat.IncludeQuotePrefix };

            if (xlCellFormat.Alignment is not null)
            {
                var alignmentKey = xlCellFormat.Alignment.ApplyTo(xlStyle.Alignment);
                xlStyle = xlStyle with { Alignment = alignmentKey };
            }

            if (xlCellFormat.Protection is not null)
            {
                var protectionKey = xlCellFormat.Protection.ApplyTo(xlStyle.Protection);
                xlStyle = xlStyle with { Protection = protectionKey };
            }

            if (xlCellFormat.NumberFormat is not null)
            {
                xlStyle = xlStyle with { NumberFormat = XLNumberFormatKey.ForFormat(xlCellFormat.NumberFormat) };
            }

            if (xlCellFormat.Font is not null)
            {
                xlStyle = xlStyle with { Font = xlCellFormat.Font.ApplyTo(xlStyle.Font) };
            }

            if (xlCellFormat.Fill is not null)
            {
                xlStyle = xlStyle with { Fill = xlCellFormat.Fill.ApplyTo(xlStyle.Fill) };
            }

            if (xlCellFormat.Border is not null)
            {
                xlStyle = xlStyle with { Border = xlCellFormat.Border.ApplyTo(xlStyle.Border) };
            }
        }

        private static XmlTreeReader CreateTreeReader(OpenXmlPart openXmlPart)
        {
            var stream = openXmlPart.GetStream(FileMode.Open);
            return new XmlTreeReader(stream, XmlToEnumMapper.Instance, false);
        }
    }
}
