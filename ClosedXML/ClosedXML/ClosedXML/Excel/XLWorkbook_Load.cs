#region

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Op = DocumentFormat.OpenXml.CustomProperties;
using Vml = DocumentFormat.OpenXml.Vml;
using Ss = DocumentFormat.OpenXml.Vml.Spreadsheet;

#endregion

namespace ClosedXML.Excel
{
    #region

    using System.Drawing;
    using Ap;
    using Op;
    using System.Xml.Linq;
    using System.Text.RegularExpressions;

    #endregion

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

        private void LoadSpreadsheetDocument(SpreadsheetDocument dSpreadsheet)
        {
            ShapeIdManager = new XLIdManager();
            SetProperties(dSpreadsheet);
            //var sharedStrings = dSpreadsheet.WorkbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>();
            SharedStringItem[] sharedStrings = null;
            if (dSpreadsheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                var shareStringPart = dSpreadsheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                sharedStrings = shareStringPart.SharedStringTable.Elements<SharedStringItem>().ToArray();
            }

            if (dSpreadsheet.CustomFilePropertiesPart != null)
            {
                foreach (var m in dSpreadsheet.CustomFilePropertiesPart.Properties.Elements<CustomDocumentProperty>())
                {
                    String name = m.Name.Value;
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
            Use1904DateSystem = wbProps != null && wbProps.Date1904 != null && wbProps.Date1904.Value;

            var wbProtection = dSpreadsheet.WorkbookPart.Workbook.WorkbookProtection;
            if (wbProtection != null)
            {
                if (wbProtection.LockStructure != null)
                    LockStructure = wbProtection.LockStructure.Value;
                if (wbProtection.LockWindows != null)
                    LockWindows = wbProtection.LockWindows.Value;
            }

            var calculationProperties = dSpreadsheet.WorkbookPart.Workbook.CalculationProperties;
            if (calculationProperties != null)
            {
                var referenceMode = calculationProperties.ReferenceMode;
                if (referenceMode != null)
                    ReferenceStyle = referenceMode.Value.ToClosedXml();

                var calculateMode = calculationProperties.CalculationMode;
                if (calculateMode != null)
                    CalculateMode = calculateMode.Value.ToClosedXml();
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
            if (s != null &&s.DifferentialFormats != null)
                differentialFormats = s.DifferentialFormats.Elements<DifferentialFormat>().ToDictionary(k => dfCount++);
            else
                differentialFormats = new Dictionary<Int32, DifferentialFormat>();
                
            var sheets = dSpreadsheet.WorkbookPart.Workbook.Sheets;
            Int32 position = 0;
            foreach (Sheet dSheet in sheets.OfType<Sheet>())
            {
                position++;
                var sharedFormulasR1C1 = new Dictionary<UInt32, String>();

                var wsPart = dSpreadsheet.WorkbookPart.GetPartById(dSheet.Id) as WorksheetPart;

                if (wsPart == null)
                {
                    UnsupportedSheets.Add(new UnsupportedSheet {SheetId = dSheet.SheetId.Value, Position = position});
                    continue;
                }

                var sheetName = dSheet.Name;

                var ws = (XLWorksheet) WorksheetsInternal.Add(sheetName, position);
                ws.RelId = dSheet.Id;
                ws.SheetId = (Int32) dSheet.SheetId.Value;


                if (dSheet.State != null)
                    ws.Visibility = dSheet.State.Value.ToClosedXml();

                var styleList = new Dictionary<int, IXLStyle>();// {{0, ws.Style}};

                using (var reader = OpenXmlReader.Create(wsPart))
                {
                    while (reader.Read())
                    {
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
                            LoadHyperlinks((Hyperlinks)reader.LoadCurrentElement(), wsPart, ws);
                        else if (reader.ElementType == typeof(PrintOptions))
                            LoadPrintOptions((PrintOptions)reader.LoadCurrentElement(), ws);
                        else if (reader.ElementType == typeof(PageMargins))
                            LoadPageMargins((PageMargins)reader.LoadCurrentElement(), ws);
                        else if (reader.ElementType == typeof(PageSetup))
                            LoadPageSetup((PageSetup)reader.LoadCurrentElement(), ws);
                        else if (reader.ElementType == typeof(HeaderFooter))
                            LoadHeaderFooter((HeaderFooter)reader.LoadCurrentElement(), ws);
                        else if (reader.ElementType == typeof(SheetProperties))
                            LoadSheetProperties((SheetProperties)reader.LoadCurrentElement(), ws);
                        else if (reader.ElementType == typeof(RowBreaks))
                            LoadRowBreaks((RowBreaks)reader.LoadCurrentElement(), ws);
                        else if (reader.ElementType == typeof(ColumnBreaks))
                            LoadColumnBreaks((ColumnBreaks)reader.LoadCurrentElement(), ws);
                        else if (reader.ElementType == typeof(LegacyDrawing))
                            ws.LegacyDrawingId = (reader.LoadCurrentElement() as LegacyDrawing).Id.Value;

                    }
                    reader.Close();
                }

                #region LoadTables

                foreach (TableDefinitionPart tablePart in wsPart.TableDefinitionParts)
                {
                    var dTable = tablePart.Table;
                    string reference = dTable.Reference.Value;
                    XLTable xlTable = ws.Range(reference).CreateTable(dTable.Name, false) as XLTable;
                    if (dTable.HeaderRowCount != null && dTable.HeaderRowCount == 0)
                    {
                        xlTable._showHeaderRow = false;
                        //foreach (var tableColumn in dTable.TableColumns.Cast<TableColumn>())
                        xlTable.AddFields(dTable.TableColumns.Cast<TableColumn>().Select(t=>GetTableColumnName(t.Name.Value)));
                    }
                    else
                    {
                        xlTable.InitializeAutoFilter();
                    }

                    if (dTable.TotalsRowCount != null && dTable.TotalsRowCount.Value > 0)
                        ((XLTable) xlTable)._showTotalsRow = true;

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
                            xlTable.Theme = (XLTableTheme) Enum.Parse(typeof (XLTableTheme), dTable.TableStyleInfo.Name.Value);
                        else
                            xlTable.Theme = XLTableTheme.None;
                    }


                    if (dTable.AutoFilter != null)
                    {
                        xlTable.ShowAutoFilter = true;
                        LoadAutoFilterColumns( dTable.AutoFilter, (xlTable as XLTable).AutoFilter);
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

                #endregion

                #region LoadComments

                if (wsPart.WorksheetCommentsPart != null) {
                    var root = wsPart.WorksheetCommentsPart.Comments;
                    var authors = root.GetFirstChild<Authors>().ChildElements;
                    var comments = root.GetFirstChild<CommentList>().ChildElements;

                    // **** MAYBE FUTURE SHAPE SIZE SUPPORT
                    XDocument xdoc = GetCommentVmlFile(wsPart);
                    
                    foreach (Comment c in comments) {
                        // find cell by reference
                        var cell = ws.Cell(c.Reference);
                        
                        XLComment xlComment = cell.Comment as XLComment;
                        xlComment.Author = authors[(int)c.AuthorId.Value].InnerText;
                        //xlComment.ShapeId = (Int32)c.ShapeId.Value;
                        //ShapeIdManager.Add(xlComment.ShapeId);

                        var runs = c.GetFirstChild<CommentText>().Elements<Run>();
                        foreach (Run run in runs) {
                            var runProperties = run.RunProperties;
                            String text = run.Text.InnerText.FixNewLines();
                            var rt = cell.Comment.AddText(text);
                            LoadFont(runProperties, rt);
                        }

                      
                        XElement shape = GetCommentShape(xdoc);
                       
                        LoadShapeProperties<IXLComment>(xlComment, shape);

                        var clientData = shape.Elements().First(e => e.Name.LocalName == "ClientData");
                        LoadClientData<IXLComment>(xlComment, clientData);

                        var textBox = shape.Elements().First(e=>e.Name.LocalName == "textbox");
                        LoadTextBox<IXLComment>(xlComment, textBox);

                        var alt = shape.Attribute("alt");
                        if (alt != null) xlComment.Style.Web.SetAlternateText(alt.Value);

                        LoadColorsAndLines<IXLComment>(xlComment, shape);

                        //var insetmode = (string)shape.Attributes().First(a=> a.Name.LocalName == "insetmode");
                        //xlComment.Style.Margins.Automatic = insetmode != null && insetmode.Equals("auto");
                        
                        shape.Remove();
                    }
                }

                #endregion
            }

            var workbook = dSpreadsheet.WorkbookPart.Workbook;

            var bookViews = workbook.BookViews;
            if (bookViews != null && bookViews.Any())
            {
                var workbookView = bookViews.First() as WorkbookView;
                if (workbookView != null && workbookView.ActiveTab != null)
                {
                    UnsupportedSheet unsupportedSheet =
                        UnsupportedSheets.FirstOrDefault(us => us.Position == (Int32) (workbookView.ActiveTab.Value + 1));
                    if (unsupportedSheet != null)
                        unsupportedSheet.IsActive = true;
                    else
                    {
                        Worksheet((Int32) (workbookView.ActiveTab.Value + 1)).SetTabActive();
                    }
                }
            }
            LoadDefinedNames(workbook);
        }

        #region Comment Helpers

        private XDocument GetCommentVmlFile(WorksheetPart wsPart)
        {
            XDocument xdoc = null;

            foreach (var vmlPart in wsPart.VmlDrawingParts)
            {
                xdoc = XDocumentExtensions.Load(vmlPart.GetStream(FileMode.Open));

                //Probe for comments
                if (xdoc.Root == null) continue;
                var shape = GetCommentShape(xdoc);
                if (shape != null) break;
            }

            if (xdoc == null) throw new Exception("Could not load comments file");
            return xdoc;
        }

        private static XElement GetCommentShape(XDocument xdoc)
        {
            var xml = xdoc.Root.Element("xml");

            XElement shape;
            if (xml != null)
                shape =
                    xml.Elements().FirstOrDefault(e => (string) e.Attribute("type") == XLConstants.Comment.ShapeTypeId);
            else
                shape = xdoc.Root.Elements().FirstOrDefault(e =>
                                                            (string) e.Attribute("type") ==
                                                            XLConstants.Comment.ShapeTypeId ||
                                                            (string) e.Attribute("type") ==
                                                            XLConstants.Comment.AlternateShapeTypeId);
            return shape;
        }

        #endregion


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

            var stroke = shape.Elements().FirstOrDefault(e=>e.Name.LocalName == "stroke");
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
                        case "single": drawing.Style.ColorsAndLines.LineStyle = XLLineStyle.Single ; break;
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
            var anchor = clientData.Elements().FirstOrDefault(e=>e.Name.LocalName == "Anchor");
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

            foreach (DefinedName definedName in workbook.DefinedNames)
            {
                var name = definedName.Name;
                if (name == "_xlnm.Print_Area")
                {
                    foreach (string area in definedName.Text.Split(','))
                    {
                        if (area.Contains("["))
                        {
                            String tableName = area.Substring(0, area.IndexOf("["));
                            var ws = Worksheets.First(w => (w as XLWorksheet).SheetId == definedName.LocalSheetId + 1);
                            ws.PageSetup.PrintAreas.Add(area);
                        }
                        else
                        {
                            string sheetName, sheetArea;
                            ParseReference(area, out sheetName, out sheetArea);
                            if (!(sheetArea.Equals("#REF") || sheetArea.EndsWith("#REF!") || sheetArea.Length == 0))
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
                            if (!NamedRanges.Any(nr => nr.Name == name))
                                NamedRanges.Add(name, text, comment);
                        }
                        else
                        {
                            if (!Worksheet(Int32.Parse(localSheetId) + 1).NamedRanges.Any(nr => nr.Name == name))
                                Worksheet(Int32.Parse(localSheetId) + 1).NamedRanges.Add(name, text, comment);
                        }
                    }
                }
            }
        }

        private void LoadPrintTitles(DefinedName definedName)
        {
            var areas = definedName.Text.Split(',');
            if (areas.Length > 0)
            {
                foreach (var item in areas)
                {
                    SetColumnsOrRowsToRepeat(item);
                }
                return;
            }

            SetColumnsOrRowsToRepeat(definedName.Text);
        }

        private void SetColumnsOrRowsToRepeat(string area)
        {
            string sheetName, sheetArea;
            ParseReference(area, out sheetName, out sheetArea);
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
            sheetName = sections[0].Replace("\'", "");
            sheetArea = sections[1];
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
                styleList.Add(styleIndex, xlCell.Style);
            }


            if (cell.CellFormula != null && cell.CellFormula.SharedIndex != null && cell.CellFormula.Reference != null)
            {
                String formula;
                if (cell.CellFormula.FormulaType != null && cell.CellFormula.FormulaType == CellFormulaValues.Array)
                    formula = "{" + cell.CellFormula.Text + "}";
                else
                    formula = cell.CellFormula.Text;

                if (cell.CellFormula.Reference != null)
                    xlCell.FormulaReference = ws.Range(cell.CellFormula.Reference.Value).RangeAddress;

                xlCell.FormulaA1 = formula;
                sharedFormulasR1C1.Add(cell.CellFormula.SharedIndex.Value, xlCell.FormulaR1C1);

                if (cell.CellValue != null)
                    xlCell.ValueCached = cell.CellValue.Text;
            }
            else if (cell.CellFormula != null)
            {
                if (cell.CellFormula.SharedIndex != null)
                    xlCell.FormulaR1C1 = sharedFormulasR1C1[cell.CellFormula.SharedIndex.Value];
                else
                {
                    String formula;
                    if (cell.CellFormula.FormulaType != null && cell.CellFormula.FormulaType == CellFormulaValues.Array)
                        formula = "{" + cell.CellFormula.Text + "}";
                    else
                        formula = cell.CellFormula.Text;

                    xlCell.FormulaA1 = formula;
                }

                if (cell.CellFormula.Reference != null)
                    xlCell.FormulaReference = ws.Range(cell.CellFormula.Reference.Value).RangeAddress;

                if (cell.CellValue != null)
                    xlCell.ValueCached = cell.CellValue.Text;
            }
            else if (cell.DataType != null)
            {
                if (cell.DataType == CellValues.InlineString)
                {
                    xlCell._cellValue = cell.InlineString != null && cell.InlineString.Text != null ? cell.InlineString.Text.Text.FixNewLines() : String.Empty;
                    xlCell._dataType = XLCellValues.Text;
                    xlCell.ShareString = false;
                }
                else if (cell.DataType == CellValues.SharedString)
                {
                    if (cell.CellValue != null)
                    {
                        if (!XLHelper.IsNullOrWhiteSpace(cell.CellValue.Text))
                        {
                            var sharedString = sharedStrings[Int32.Parse(cell.CellValue.Text)];

                            var runs = sharedString.Elements<Run>();
                            var phoneticRuns = sharedString.Elements<PhoneticRun>();
                            var phoneticProperties = sharedString.Elements<PhoneticProperties>();
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
                        
                            if(!hasRuns)
                                xlCell._cellValue = sharedString.Text.InnerText;

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

                            #endregion

                            #region Load Phonetic Runs

                            foreach (PhoneticRun pr in phoneticRuns)
                            {
                                xlCell.RichText.Phonetics.Add(pr.Text.InnerText.FixNewLines(), (Int32)pr.BaseTextStartIndex.Value,
                                                              (Int32) pr.EndingBaseIndex.Value);
                            }

                            #endregion
                        }
                        else
                            xlCell._cellValue = cell.CellValue.Text.FixNewLines();
                    }
                    else
                        xlCell._cellValue = String.Empty;
                    xlCell._dataType = XLCellValues.Text;
                }
                else if (cell.DataType == CellValues.Date)
                {
                    if (!XLHelper.IsNullOrWhiteSpace(cell.CellValue.Text))
                        xlCell._cellValue = Double.Parse(cell.CellValue.Text, CultureInfo.InvariantCulture).ToString();
                    xlCell._dataType = XLCellValues.DateTime;
                }
                else if (cell.DataType == CellValues.Boolean)
                {
                    xlCell._cellValue = cell.CellValue.Text;
                    xlCell._dataType = XLCellValues.Boolean;
                }
                else if (cell.DataType == CellValues.Number)
                {
                    if (!XLHelper.IsNullOrWhiteSpace(cell.CellValue.Text))
                        xlCell._cellValue = Double.Parse(cell.CellValue.Text, CultureInfo.InvariantCulture).ToString();
                    if (s == null)
                    {
                        xlCell._dataType = XLCellValues.Number;
                    }
                    else
                    {
                        var numberFormatId = ((CellFormat)(s.CellFormats).ElementAt(styleIndex)).NumberFormatId;
                        if (numberFormatId == 46U)
                            xlCell.DataType = XLCellValues.TimeSpan;
                        else
                            xlCell._dataType = XLCellValues.Number;
                    }
                    
                }
            }
            else if (cell.CellValue != null)
            {
                if (s == null)
                {
                    xlCell._dataType = XLCellValues.Number;
                }
                else
                {
                    var numberFormatId = ((CellFormat) (s.CellFormats).ElementAt(styleIndex)).NumberFormatId;
                    if (!XLHelper.IsNullOrWhiteSpace(cell.CellValue.Text))
                        xlCell._cellValue = Double.Parse(cell.CellValue.Text, CultureInfo.InvariantCulture).ToString();
                    if (s.NumberingFormats != null &&
                        s.NumberingFormats.Any(nf => ((NumberingFormat) nf).NumberFormatId.Value == numberFormatId))
                    {
                        xlCell.Style.NumberFormat.Format =
                            ((NumberingFormat) s.NumberingFormats
                                                .First(
                                                    nf => ((NumberingFormat) nf).NumberFormatId.Value == numberFormatId)
                            ).FormatCode.Value;
                    }
                    else
                        xlCell.Style.NumberFormat.NumberFormatId = Int32.Parse(numberFormatId);


                    if (!XLHelper.IsNullOrWhiteSpace(xlCell.Style.NumberFormat.Format))
                        xlCell._dataType = GetDataTypeFromFormat(xlCell.Style.NumberFormat.Format);
                    else if ((numberFormatId >= 14 && numberFormatId <= 22) ||
                             (numberFormatId >= 45 && numberFormatId <= 47))
                        xlCell._dataType = XLCellValues.DateTime;
                    else if (numberFormatId == 49)
                        xlCell._dataType = XLCellValues.Text;
                    else
                        xlCell._dataType = XLCellValues.Number;
                }
            }
        }

        private void LoadNumberFormat(NumberingFormat nfSource, IXLNumberFormat nf)
        {
            if (nfSource == null) return;

            if (nfSource.FormatCode != null)
                nf.Format = nfSource.FormatCode.Value;
            //if (nfSource.NumberFormatId != null)
            //    nf.NumberFormatId = (Int32)nfSource.NumberFormatId.Value;
        }

        private void LoadBorder(Border borderSource, IXLBorder border)
        {
            if (borderSource == null) return;

            LoadBorderValues(borderSource.DiagonalBorder, border.SetDiagonalBorder, border.SetDiagonalBorderColor);

            if (borderSource.DiagonalUp != null )
                border.DiagonalUp = borderSource.DiagonalUp.Value;
            if (borderSource.DiagonalDown != null)
                border.DiagonalDown = borderSource.DiagonalDown.Value;

            LoadBorderValues(borderSource.LeftBorder, border.SetLeftBorder, border.SetLeftBorderColor);
            LoadBorderValues(borderSource.RightBorder, border.SetRightBorder, border.SetRightBorderColor);
            LoadBorderValues(borderSource.TopBorder, border.SetTopBorder, border.SetTopBorderColor);
            LoadBorderValues(borderSource.BottomBorder, border.SetBottomBorder, border.SetBottomBorderColor);
            
        }

        private void LoadBorderValues(BorderPropertiesType source, Func<XLBorderStyleValues, IXLStyle> setBorder, Func<XLColor, IXLStyle> setColor )
        {
            if (source != null)
            {
                if (source.Style != null)
                    setBorder(source.Style.Value.ToClosedXml());
                if (source.Color != null)
                    setColor(GetColor(source.Color));
            }
        }

        

        private void LoadFill(Fill fillSource, IXLFill fill)
        {
            if (fillSource == null) return;

            if(fillSource.PatternFill != null)
            {
                if (fillSource.PatternFill.PatternType != null)
                    fill.PatternType = fillSource.PatternFill.PatternType.Value.ToClosedXml();
                else
                    fill.PatternType = XLFillPatternValues.Solid;

                if (fillSource.PatternFill.ForegroundColor != null)
                    fill.PatternColor = GetColor(fillSource.PatternFill.ForegroundColor);
                if (fillSource.PatternFill.BackgroundColor != null)
                    fill.PatternBackgroundColor = GetColor(fillSource.PatternFill.BackgroundColor);
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
                    (XLFontFamilyNumberingValues) Int32.Parse(fontFamilyNumbering.Val.ToString());
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
            Int32 rowIndex = row.RowIndex == null ? ++lastRow : (Int32) row.RowIndex.Value;
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
                columns.Elements<Column>().Where(c => c.Max == XLHelper.MaxColumnNumber).FirstOrDefault();

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

                var xlColumns = (XLColumns) ws.Columns(col.Min, col.Max);
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

        private static XLCellValues GetDataTypeFromFormat(String format)
        {
            int length = format.Length;
            String f = format.ToLower();
            for (Int32 i = 0; i < length; i++)
            {
                Char c = f[i];
                if (c == '"')
                    i = f.IndexOf('"', i + 1);
                else if (c == '0' || c == '#' || c == '?')
                    return XLCellValues.Number;
                else if (c == 'y' || c == 'm' || c == 'd' || c == 'h' || c == 's')
                    return XLCellValues.DateTime;
            }
            return XLCellValues.Text;
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
                    foreach (CustomFilter filter in filterColumn.CustomFilters)
                    {
                        Double dTest;
                        String val = filter.Val.Value;
                        if (!Double.TryParse(val, out dTest))
                        {
                            isText = true;
                            break;
                        }
                    }

                    foreach (CustomFilter filter in filterColumn.CustomFilters)
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
                                    condition = o => o.ToString().Equals(xlFilter.Value.ToString(), StringComparison.InvariantCultureIgnoreCase);
                                else
                                    condition = o => (o as IComparable).CompareTo(xlFilter.Value) == 0;
                                break;
                            case XLFilterOperator.EqualOrGreaterThan: condition = o => (o as IComparable).CompareTo(xlFilter.Value) >= 0; break;
                            case XLFilterOperator.EqualOrLessThan: condition = o => (o as IComparable).CompareTo(xlFilter.Value) <= 0; break;
                            case XLFilterOperator.GreaterThan: condition = o => (o as IComparable).CompareTo(xlFilter.Value) > 0; break;
                            case XLFilterOperator.LessThan: condition = o => (o as IComparable).CompareTo(xlFilter.Value) < 0; break;
                            case XLFilterOperator.NotEqual:
                                if (isText)
                                    condition = o => !o.ToString().Equals(xlFilter.Value.ToString(), StringComparison.InvariantCultureIgnoreCase);
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
                    var filterList = new List<XLFilter>();
                    autoFilter.Column(column).FilterType = XLFilterType.Regular;
                    autoFilter.Filters.Add((int)filterColumn.ColumnId.Value + 1, filterList);

                    Boolean isText = false;
                    foreach (Filter filter in filterColumn.Filters.OfType<Filter>())
                    {
                        Double dTest;
                        String val = filter.Val.Value;
                        if (!Double.TryParse(val, out dTest))
                        {
                            isText = true;
                            break;
                        }
                    }

                    foreach (Filter filter in filterColumn.Filters.OfType<Filter>())
                    {
                        var xlFilter = new XLFilter { Connector = XLConnector.Or, Operator = XLFilterOperator.Equal };

                        Func<Object, Boolean> condition;
                        if (isText)
                        {
                            xlFilter.Value = filter.Val.Value;
                            condition = o => o.ToString().Equals(xlFilter.Value.ToString(), StringComparison.InvariantCultureIgnoreCase);
                        }
                        else
                        {
                            xlFilter.Value = Double.Parse(filter.Val.Value, CultureInfo.InvariantCulture);
                            condition = o => (o as IComparable).CompareTo(xlFilter.Value) == 0;
                        }

                        xlFilter.Condition = condition;
                        filterList.Add(xlFilter);
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
                    Int32 column = ws.Range(condition.Reference.Value).FirstCell().Address.ColumnNumber - autoFilter.Range.FirstCell().Address.ColumnNumber + 1 ;
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
            if (sp.SelectLockedCells != null) ws.Protection.SelectLockedCells = sp.SelectLockedCells.Value;
            if (sp.SelectUnlockedCells != null) ws.Protection.SelectUnlockedCells = sp.SelectUnlockedCells.Value;
        }

        private static void LoadDataValidations(DataValidations dataValidations, XLWorksheet ws)
        {
            if (dataValidations == null) return;

            foreach (DataValidation dvs in dataValidations.Elements<DataValidation>())
            {
                String txt = dvs.SequenceOfReferences.InnerText;
                if (XLHelper.IsNullOrWhiteSpace(txt)) continue;
                foreach (var dvt in txt.Split(' ').Select(rangeAddress => ws.Range(rangeAddress).DataValidation))
                {
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

        private void LoadConditionalFormatting(ConditionalFormatting conditionalFormatting, XLWorksheet ws, Dictionary<Int32, DifferentialFormat> differentialFormats)
        {
            if (conditionalFormatting == null) return;

            foreach (var sor in conditionalFormatting.SequenceOfReferences.Items)
            {
                foreach (var fr in conditionalFormatting.Elements<ConditionalFormattingRule>())
                {
                    var conditionalFormat = new XLConditionalFormat(ws.Range(sor.Value));
                    if (fr.FormatId != null)
                    {
                        LoadFont(differentialFormats[(Int32) fr.FormatId.Value].Font, conditionalFormat.Style.Font);
                        LoadFill(differentialFormats[(Int32) fr.FormatId.Value].Fill, conditionalFormat.Style.Fill);
                        LoadBorder(differentialFormats[(Int32) fr.FormatId.Value].Border, conditionalFormat.Style.Border);
                        LoadNumberFormat(differentialFormats[(Int32) fr.FormatId.Value].NumberingFormat, conditionalFormat.Style.NumberFormat);
                    }
                    if (fr.Operator != null)
                        conditionalFormat.Operator = fr.Operator.Value.ToClosedXml();
                    if (fr.Type != null)
                        conditionalFormat.ConditionalFormatType = fr.Type.Value.ToClosedXml();
                    if (fr.Text != null)
                        conditionalFormat.Values.Add(GetFormula(fr.Text.Value));
                    if (fr.Percent != null)
                        conditionalFormat.Percent = fr.Percent.Value;
                    if (fr.Bottom != null)
                        conditionalFormat.Bottom = fr.Bottom.Value;
                    if (fr.Rank != null)
                        conditionalFormat.Values.Add(GetFormula(fr.Rank.Value.ToString()));

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

        private void LoadSheetProperties(SheetProperties sheetProperty, XLWorksheet ws)
        {
            if (sheetProperty == null) return;

            if (sheetProperty.TabColor != null)
                ws.TabColor = GetColor(sheetProperty.TabColor);

            if (sheetProperty.OutlineProperties == null) return;

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
            var xlFooter = (XLHeaderFooter) ws.PageSetup.Footer;
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
            var xlHeader = (XLHeaderFooter) ws.PageSetup.Header;
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

        private static void LoadPageSetup(PageSetup pageSetup, XLWorksheet ws)
        {
            if (pageSetup == null) return;

            if (pageSetup.PaperSize != null)
                ws.PageSetup.PaperSize = (XLPaperSize) Int32.Parse(pageSetup.PaperSize.InnerText);
            if (pageSetup.Scale != null)
                ws.PageSetup.Scale = Int32.Parse(pageSetup.Scale.InnerText);
            else
            {
                if (pageSetup.FitToWidth != null)
                    ws.PageSetup.PagesWide = Int32.Parse(pageSetup.FitToWidth.InnerText);
                if (pageSetup.FitToHeight != null)
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
            if (pageSetup.HorizontalDpi != null) ws.PageSetup.HorizontalDpi = (Int32) pageSetup.HorizontalDpi.Value;
            if (pageSetup.VerticalDpi != null) ws.PageSetup.VerticalDpi = (Int32) pageSetup.VerticalDpi.Value;
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

            var pane = sheetView.Elements<Pane>().FirstOrDefault();

            if (pane == null) return;

            if (pane.State == null ||
                (pane.State != PaneStateValues.FrozenSplit && pane.State != PaneStateValues.Frozen)) return;

            if (pane.HorizontalSplit != null)
                ws.SheetView.SplitColumn = (Int32) pane.HorizontalSplit.Value;
            if (pane.VerticalSplit != null)
                ws.SheetView.SplitRow = (Int32) pane.VerticalSplit.Value;

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
                        thisColor = ColorTranslator.FromHtml(htmlColor);
                        _colorList.Add(htmlColor, thisColor);
                    }
                    else
                        thisColor = _colorList[htmlColor];
                    retVal = XLColor.FromColor(thisColor);
                }
                else if (color.Indexed != null && color.Indexed < 64)
                    retVal = XLColor.FromIndex((Int32) color.Indexed.Value);
                else if (color.Theme != null)
                {
                    retVal = color.Tint != null ? XLColor.FromTheme((XLThemeColor) color.Theme.Value, color.Tint.Value) : XLColor.FromTheme((XLThemeColor) color.Theme.Value);
                }
            }
            return retVal ?? XLColor.NoColor;
        }

        private void ApplyStyle(IXLStylized xlStylized, Int32 styleIndex, Stylesheet s, Fills fills, Borders borders,
                                Fonts fonts, NumberingFormats numberingFormats)
        {
            if (s == null) return; //No Stylesheet, no Styles

            var cellFormat = (CellFormat) s.CellFormats.ElementAt(styleIndex);

            if (cellFormat.ApplyProtection != null)
            {
                var protection = cellFormat.Protection;

                if (protection == null)
                    xlStylized.InnerStyle.Protection = new XLProtection(null, DefaultStyle.Protection);
                else
                {
                    xlStylized.InnerStyle.Protection.Hidden = protection.Hidden != null && protection.Hidden.HasValue &&
                                                              protection.Hidden.Value;
                    xlStylized.InnerStyle.Protection.Locked = protection.Locked == null ||
                                                              (protection.Locked.HasValue && protection.Locked.Value);
                }
            }

            if (UInt32HasValue(cellFormat.FillId))
            {
                var fill = (Fill)fills.ElementAt((Int32)cellFormat.FillId.Value);
                if (fill.PatternFill != null)
                {
                    if (fill.PatternFill.PatternType != null)
                        xlStylized.InnerStyle.Fill.PatternType = fill.PatternFill.PatternType.Value.ToClosedXml();

                    var fgColor = GetColor(fill.PatternFill.ForegroundColor);
                    if (fgColor.HasValue) xlStylized.InnerStyle.Fill.PatternColor = fgColor;

                    var bgColor = GetColor(fill.PatternFill.BackgroundColor);
                    if (bgColor.HasValue)
                        xlStylized.InnerStyle.Fill.PatternBackgroundColor = bgColor;
                }
            }


            var alignment = cellFormat.Alignment;
            if (alignment != null)
            {
                if (alignment.Horizontal != null)
                    xlStylized.InnerStyle.Alignment.Horizontal = alignment.Horizontal.Value.ToClosedXml();
                if (alignment.Indent != null && alignment.Indent != 0)
                    xlStylized.InnerStyle.Alignment.Indent = Int32.Parse(alignment.Indent.ToString());
                if (alignment.JustifyLastLine != null)
                    xlStylized.InnerStyle.Alignment.JustifyLastLine = alignment.JustifyLastLine;
                if (alignment.ReadingOrder != null)
                {
                    xlStylized.InnerStyle.Alignment.ReadingOrder =
                        (XLAlignmentReadingOrderValues) Int32.Parse(alignment.ReadingOrder.ToString());
                }
                if (alignment.RelativeIndent != null)
                    xlStylized.InnerStyle.Alignment.RelativeIndent = alignment.RelativeIndent;
                if (alignment.ShrinkToFit != null)
                    xlStylized.InnerStyle.Alignment.ShrinkToFit = alignment.ShrinkToFit;
                if (alignment.TextRotation != null)
                    xlStylized.InnerStyle.Alignment.TextRotation = (Int32) alignment.TextRotation.Value;
                if (alignment.Vertical != null)
                    xlStylized.InnerStyle.Alignment.Vertical = alignment.Vertical.Value.ToClosedXml();
                if (alignment.WrapText != null)
                    xlStylized.InnerStyle.Alignment.WrapText = alignment.WrapText;
            }


            if (UInt32HasValue(cellFormat.BorderId))
            {
                uint borderId = cellFormat.BorderId.Value;
                var border = (Border)borders.ElementAt((Int32)borderId);
                if (border != null)
                {
                    var bottomBorder = border.BottomBorder;
                    if (bottomBorder != null)
                    {
                        if (bottomBorder.Style != null)
                            xlStylized.InnerStyle.Border.BottomBorder = bottomBorder.Style.Value.ToClosedXml();

                        var bottomBorderColor = GetColor(bottomBorder.Color);
                        if (bottomBorderColor.HasValue)
                            xlStylized.InnerStyle.Border.BottomBorderColor = bottomBorderColor;
                    }
                    var topBorder = border.TopBorder;
                    if (topBorder != null)
                    {
                        if (topBorder.Style != null)
                            xlStylized.InnerStyle.Border.TopBorder = topBorder.Style.Value.ToClosedXml();
                        var topBorderColor = GetColor(topBorder.Color);
                        if (topBorderColor.HasValue)
                            xlStylized.InnerStyle.Border.TopBorderColor = topBorderColor;
                    }
                    var leftBorder = border.LeftBorder;
                    if (leftBorder != null)
                    {
                        if (leftBorder.Style != null)
                            xlStylized.InnerStyle.Border.LeftBorder = leftBorder.Style.Value.ToClosedXml();
                        var leftBorderColor = GetColor(leftBorder.Color);
                        if (leftBorderColor.HasValue)
                            xlStylized.InnerStyle.Border.LeftBorderColor = leftBorderColor;
                    }
                    var rightBorder = border.RightBorder;
                    if (rightBorder != null)
                    {
                        if (rightBorder.Style != null)
                            xlStylized.InnerStyle.Border.RightBorder = rightBorder.Style.Value.ToClosedXml();
                        var rightBorderColor = GetColor(rightBorder.Color);
                        if (rightBorderColor.HasValue)
                            xlStylized.InnerStyle.Border.RightBorderColor = rightBorderColor;
                    }
                    var diagonalBorder = border.DiagonalBorder;
                    if (diagonalBorder != null)
                    {
                        if (diagonalBorder.Style != null)
                            xlStylized.InnerStyle.Border.DiagonalBorder = diagonalBorder.Style.Value.ToClosedXml();
                        var diagonalBorderColor = GetColor(diagonalBorder.Color);
                        if (diagonalBorderColor.HasValue)
                            xlStylized.InnerStyle.Border.DiagonalBorderColor = diagonalBorderColor;
                        if (border.DiagonalDown != null)
                            xlStylized.InnerStyle.Border.DiagonalDown = border.DiagonalDown;
                        if (border.DiagonalUp != null)
                            xlStylized.InnerStyle.Border.DiagonalUp = border.DiagonalUp;
                    }
                }
            }

            if (UInt32HasValue(cellFormat.FontId))
            {
                var fontId = cellFormat.FontId;
                var font = (DocumentFormat.OpenXml.Spreadsheet.Font)fonts.ElementAt((Int32)fontId.Value);
                if (font != null)
                {
                    xlStylized.InnerStyle.Font.Bold = GetBoolean(font.Bold);

                    var fontColor = GetColor(font.Color);
                    if (fontColor.HasValue)
                        xlStylized.InnerStyle.Font.FontColor = fontColor;

                    if (font.FontFamilyNumbering != null && (font.FontFamilyNumbering).Val != null)
                    {
                        xlStylized.InnerStyle.Font.FontFamilyNumbering =
                            (XLFontFamilyNumberingValues)Int32.Parse((font.FontFamilyNumbering).Val.ToString());
                    }
                    if (font.FontName != null)
                    {
                        if ((font.FontName).Val != null)
                            xlStylized.InnerStyle.Font.FontName = (font.FontName).Val;
                    }
                    if (font.FontSize != null)
                    {
                        if ((font.FontSize).Val != null)
                            xlStylized.InnerStyle.Font.FontSize = (font.FontSize).Val;
                    }

                    xlStylized.InnerStyle.Font.Italic = GetBoolean(font.Italic);
                    xlStylized.InnerStyle.Font.Shadow = GetBoolean(font.Shadow);
                    xlStylized.InnerStyle.Font.Strikethrough = GetBoolean(font.Strike);

                    if (font.Underline != null)
                    {
                        xlStylized.InnerStyle.Font.Underline = font.Underline.Val != null
                                                                   ? (font.Underline).Val.Value.ToClosedXml()
                                                                   : XLFontUnderlineValues.Single;
                    }

                    if (font.VerticalTextAlignment != null)
                    {
                        xlStylized.InnerStyle.Font.VerticalAlignment = font.VerticalTextAlignment.Val != null
                                                                           ? (font.VerticalTextAlignment).Val.Value.
                                                                                 ToClosedXml()
                                                                           : XLFontVerticalTextAlignmentValues.Baseline;
                    }
                }
            }

            

            if (!UInt32HasValue(cellFormat.NumberFormatId)) return;
            
            var numberFormatId = cellFormat.NumberFormatId;

            string formatCode = String.Empty;
            if (numberingFormats != null)
            {
                var numberingFormat =
                    numberingFormats.FirstOrDefault(
                        nf =>
                        ((NumberingFormat) nf).NumberFormatId != null &&
                        ((NumberingFormat) nf).NumberFormatId.Value == numberFormatId) as NumberingFormat;

                if (numberingFormat != null && numberingFormat.FormatCode != null)
                    formatCode = numberingFormat.FormatCode.Value;
            }
            if (formatCode.Length > 0)
                xlStylized.InnerStyle.NumberFormat.Format = formatCode;
            else
                xlStylized.InnerStyle.NumberFormat.NumberFormatId = (Int32) numberFormatId.Value;
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