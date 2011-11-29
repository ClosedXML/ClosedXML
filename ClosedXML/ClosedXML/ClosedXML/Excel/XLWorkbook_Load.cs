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
                                             DateTime.ParseExact(m.VTFileTime.Text, "yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'",
                                                                 CultureInfo.InvariantCulture));
                    }
                    else if (m.VTDouble != null)
                        CustomProperties.Add(name, Double.Parse(m.VTDouble.Text, CultureInfo.InvariantCulture));
                    else if (m.VTBool != null)
                        CustomProperties.Add(name, m.VTBool.Text == "true");
                }
            }

            var referenceMode = dSpreadsheet.WorkbookPart.Workbook.CalculationProperties.ReferenceMode;
            if (referenceMode != null)
                ReferenceStyle = referenceMode.Value.ToClosedXml();

            var calculateMode = dSpreadsheet.WorkbookPart.Workbook.CalculationProperties.CalculationMode;
            if (calculateMode != null)
                CalculateMode = calculateMode.Value.ToClosedXml();

            if (dSpreadsheet.ExtendedFilePropertiesPart.Properties.Elements<Company>().Count() > 0)
                Properties.Company = dSpreadsheet.ExtendedFilePropertiesPart.Properties.GetFirstChild<Company>().Text;

            if (dSpreadsheet.ExtendedFilePropertiesPart.Properties.Elements<Manager>().Count() > 0)
                Properties.Manager = dSpreadsheet.ExtendedFilePropertiesPart.Properties.GetFirstChild<Manager>().Text;


            var workbookStylesPart = dSpreadsheet.WorkbookPart.WorkbookStylesPart;
            var s = workbookStylesPart.Stylesheet;

            var numberingFormats = s.NumberingFormats;
            //Int32 fillCount = (Int32)s.Fills.Count.Value;
            var fills = s.Fills;
            var borders = s.Borders;
            var fonts = s.Fonts;

            var sheets = dSpreadsheet.WorkbookPart.Workbook.Sheets;
            Int32 position = 0;
            foreach (Sheet dSheet in sheets.OfType<Sheet>())
            {
                position++;
                var sharedFormulasR1C1 = new Dictionary<UInt32, String>();

                var wsPart = dSpreadsheet.WorkbookPart.GetPartById(dSheet.Id) as WorksheetPart;

                if (wsPart == null)
                {
                    UnsupportedSheets.Add(position, new UnsupportedSheet {SheetId = dSheet.SheetId.Value});
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
                                    ws.Range(mergeCell.Reference).Merge();
                            }
                        }
                        else if (reader.ElementType == typeof(Columns))
                            LoadColumns(s, numberingFormats, fills, borders, fonts, ws,
                                        (Columns)reader.LoadCurrentElement());
                        else if (reader.ElementType == typeof(Row))
                        {
                            LoadRows(s, numberingFormats, fills, borders, fonts, ws, sharedStrings, sharedFormulasR1C1,
                                     styleList, (Row)reader.LoadCurrentElement());
                        }
                        else if (reader.ElementType == typeof(AutoFilter))
                            LoadAutoFilter((AutoFilter)reader.LoadCurrentElement(), ws);
                        else if (reader.ElementType == typeof(SheetProtection))
                            LoadSheetProtection((SheetProtection)reader.LoadCurrentElement(), ws);
                        else if (reader.ElementType == typeof(DataValidations))
                            LoadDataValidations((DataValidations)reader.LoadCurrentElement(), ws);
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

                    }
                    reader.Close();
                }

                #region LoadTables

                foreach (TableDefinitionPart tablePart in wsPart.TableDefinitionParts)
                {
                    var dTable = tablePart.Table;
                    string reference = dTable.Reference.Value;
                    var xlTable = ws.Range(reference).CreateTable(dTable.Name);
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
                            xlTable.Theme =
                                (XLTableTheme) Enum.Parse(typeof (XLTableTheme), dTable.TableStyleInfo.Name.Value);
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
                            if (tableColumn.TotalsRowFunction != null)
                                xlTable.Field(tableColumn.Name.Value).TotalsRowFunction =
                                    tableColumn.TotalsRowFunction.Value.ToClosedXml();

                            if (tableColumn.TotalsRowFormula != null)
                                xlTable.Field(tableColumn.Name.Value).TotalsRowFormulaA1 =
                                    tableColumn.TotalsRowFormula.Text;

                            if (tableColumn.TotalsRowLabel != null)
                                xlTable.Field(tableColumn.Name.Value).TotalsRowLabel = tableColumn.TotalsRowLabel.Value;
                        }

                        xlTable.AutoFilter.Range = xlTable.Worksheet.Range(
                                                    xlTable.RangeAddress.FirstAddress.RowNumber, xlTable.RangeAddress.FirstAddress.ColumnNumber,
                                                    xlTable.RangeAddress.LastAddress.RowNumber - 1, xlTable.RangeAddress.LastAddress.ColumnNumber);
                    }
                    else
                        xlTable.AutoFilter.Range = xlTable.Worksheet.Range(xlTable.RangeAddress);
                }

                #endregion

                #region LoadComments

                if (wsPart.WorksheetCommentsPart != null) {
                    var root = wsPart.WorksheetCommentsPart.Comments;
                    var authors = root.GetFirstChild<Authors>().ChildElements;
                    var comments = root.GetFirstChild<CommentList>().ChildElements;

                    // **** MAYBE FUTURE SHAPE SIZE SUPPORT
                    // var shapes = wsPart.VmlDrawingParts.SelectMany(p => new System.Xml.XmlTextReader(p.GetStream()).Read()

                    foreach (Comment c in comments) {
                        // find cell by reference
                        var cell = ws.Cell(c.Reference);
                        cell.Comment.Author = authors[(int)c.AuthorId.Value].InnerText;
                        var runs = c.GetFirstChild<CommentText>().Elements<Run>();
                        foreach (Run run in runs) {
                            var runProperties = run.RunProperties;
                            String text = run.Text.InnerText.FixNewLines();
                            var rt = cell.Comment.AddText(text);
                            LoadFont(runProperties, rt);
                        }

                        // **** MAYBE FUTURE SHAPE SIZE SUPPORT
                        //var shape = shapes.FirstOrDefault(sh => { 
                        //                        var cd = sh.GetFirstChild<Ss.ClientData>(); 
                        //                        return cd.GetFirstChild<Ss.CommentRowTarget>().InnerText == cell.Address.RowNumber.ToString()
                        //                            && cd.GetFirstChild<Ss.CommentColumnTarget>().InnerText == cell.Address.ColumnNumber.ToString();
                        //                        });

                        //var location = shape.GetFirstChild<Ss.Anchor>().InnerText.Split(',');

                        //var leftCol = int.Parse(location[0]);
                        //var leftOffsetPx = int.Parse(location[1]);
                        //var topRow = int.Parse(location[2]);
                        //var topOffsetPx = int.Parse(location[3]);
                        //var rightCol = int.Parse(location[4]);
                        //var riightOffsetPx = int.Parse(location[5]);
                        //var bottomRow = int.Parse(location[6]);
                        //var bottomOffsetPx = int.Parse(location[7]);

                        //cmt.Style.Size.Height = bottomRow - topRow;
                        //cmt.Style.Size.Width = rightCol = leftCol;

                    }

                }

                #endregion
            }

            var workbook = dSpreadsheet.WorkbookPart.Workbook;

            var workbookView = (WorkbookView) workbook.BookViews.FirstOrDefault();
            if (workbookView != null && workbookView.ActiveTab != null)
            {
                UnsupportedSheet unsupportedSheet;
                if (UnsupportedSheets.TryGetValue((Int32)(workbookView.ActiveTab.Value + 1), out unsupportedSheet))
                    unsupportedSheet.IsActive = true;
                else
                {
                    Int32 sId = (Int32)(workbookView.ActiveTab.Value + 1);
                    Worksheet(sId).SetTabActive();
                    //- _unsupportedSheets.Keys.Where(n=>n <= sId ).Count()
                }
            }

            LoadDefinedNames(workbook);
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
                        string sheetName, sheetArea;
                        ParseReference(area, out sheetName, out sheetArea);
                        if (!(sheetArea.Equals("#REF") || sheetArea.EndsWith("#REF!")))
                            WorksheetsInternal.Worksheet(sheetName).PageSetup.PrintAreas.Add(sheetArea);
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

        private void LoadCells(SharedStringItem[] sharedStrings, Stylesheet s, NumberingFormats numberingFormats,
                               Fills fills, Borders borders, Fonts fonts, Dictionary<uint, string> sharedFormulasR1C1,
                               XLWorksheet ws, Dictionary<Int32, IXLStyle> styleList, Cell cell)
        {
            Int32 styleIndex = cell.StyleIndex != null ? Int32.Parse(cell.StyleIndex.InnerText) : 0;
            var xlCell = ws.CellFast(cell.CellReference);

            if (styleList.ContainsKey(styleIndex))
                xlCell.Style = styleList[styleIndex];
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

                if (cell.CellValue != null)
                    xlCell.ValueCached = cell.CellValue.Text;
            }
            else if (cell.DataType != null)
            {
                if (cell.DataType == CellValues.InlineString)
                {
                    xlCell._cellValue = cell.InlineString.Text.Text.FixNewLines(); 
                    xlCell._dataType = XLCellValues.Text;
                    xlCell.ShareString = false;
                }
                else if (cell.DataType == CellValues.SharedString)
                {
                    if (cell.CellValue != null)
                    {
                        if (!StringExtensions.IsNullOrWhiteSpace(cell.CellValue.Text))
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
                    //xlCell.cellValue = DateTime.FromOADate(Double.Parse(dCell.CellValue.Text, CultureInfo.InvariantCulture));
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
                    xlCell._cellValue = Double.Parse(cell.CellValue.Text, CultureInfo.InvariantCulture).ToString();
                    var numberFormatId = ((CellFormat) (s.CellFormats).ElementAt(styleIndex)).NumberFormatId;
                    if (numberFormatId == 46U)
                        xlCell.DataType = XLCellValues.TimeSpan;
                    else
                        xlCell._dataType = XLCellValues.Number;
                }
            }
            else if (cell.CellValue != null)
            {
                var numberFormatId = ((CellFormat) (s.CellFormats).ElementAt(styleIndex)).NumberFormatId;
                xlCell._cellValue = Double.Parse(cell.CellValue.Text, CultureInfo.InvariantCulture).ToString();
                if (s.NumberingFormats != null &&
                    s.NumberingFormats.Any(nf => ((NumberingFormat) nf).NumberFormatId.Value == numberFormatId))
                {
                    xlCell.Style.NumberFormat.Format =
                        ((NumberingFormat)
                         s.NumberingFormats.Where(nf => ((NumberingFormat) nf).NumberFormatId.Value == numberFormatId).
                             Single()).FormatCode.Value;
                }
                else
                    xlCell.Style.NumberFormat.NumberFormatId = Int32.Parse(numberFormatId);


                if (!StringExtensions.IsNullOrWhiteSpace(xlCell.Style.NumberFormat.Format))
                    xlCell._dataType = GetDataTypeFromFormat(xlCell.Style.NumberFormat.Format);
                else if ((numberFormatId >= 14 && numberFormatId <= 22) || (numberFormatId >= 45 && numberFormatId <= 47))
                    xlCell._dataType = XLCellValues.DateTime;
                else if (numberFormatId == 49)
                    xlCell._dataType = XLCellValues.Text;
                else
                    xlCell._dataType = XLCellValues.Number;
            }
        }

        private void LoadFont(OpenXmlElement fontSource, IXLFontBase fontBase)
        {
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

        private void LoadRows(Stylesheet s, NumberingFormats numberingFormats, Fills fills, Borders borders, Fonts fonts,
                              XLWorksheet ws, SharedStringItem[] sharedStrings,
                              Dictionary<uint, string> sharedFormulasR1C1, Dictionary<Int32, IXLStyle> styleList,
                              Row row)
        {
            var xlRow = ws.Row((Int32) row.RowIndex.Value, false);
            if (row.Height != null)
                xlRow.Height = row.Height;
            else
                xlRow.Height = ws.RowHeight;

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
                    ApplyStyle(xlRow, styleIndex, s, fills, borders, fonts, numberingFormats);
                else
                {
                    xlRow.Style = DefaultStyle;
                }
            }

            foreach (Cell cell in row.Elements<Cell>())
                LoadCells(sharedStrings, s, numberingFormats, fills, borders, fonts, sharedFormulasR1C1, ws, styleList,
                          cell);
        }

        private void LoadColumns(Stylesheet s, NumberingFormats numberingFormats, Fills fills, Borders borders,
                                 Fonts fonts, XLWorksheet ws, Columns columns)
        {
            if (columns == null) return;

            var wsDefaultColumn =
                columns.Elements<Column>().Where(c => c.Max == ExcelHelper.MaxColumnNumber).FirstOrDefault();

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
                if (col.Max == ExcelHelper.MaxColumnNumber) continue;

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
                    ApplyStyle(xlColumns, styleIndex, s, fills, borders, fonts, numberingFormats);
                else
                    xlColumns.Style = DefaultStyle;
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
                            xlFilter.Value = Double.Parse(filter.Val.Value);

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
                    foreach (Filter filter in filterColumn.Filters)
                    {
                        Double dTest;
                        String val = filter.Val.Value;
                        if (!Double.TryParse(val, out dTest))
                        {
                            isText = true;
                            break;
                        }
                    }

                    foreach (Filter filter in filterColumn.Filters)
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
                            xlFilter.Value = Double.Parse(filter.Val.Value);
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
                foreach (var dvt in dvs.SequenceOfReferences.InnerText.Split(' ').Select(rangeAddress => ws.Range(rangeAddress).DataValidation))
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
                    xlCell.Hyperlink = hl.Id != null ? new XLHyperlink(hyperlinkDictionary[hl.Id], tooltip) : new XLHyperlink(hl.Location.Value, tooltip);
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
                ws.PageSetup.FirstPageNumber = Int32.Parse(pageSetup.FirstPageNumber.InnerText);
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

        private IXLColor GetColor(ColorType color)
        {
            IXLColor retVal = null;
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
                    retVal = new XLColor(thisColor);
                }
                else if (color.Indexed != null && color.Indexed < 64)
                    retVal = new XLColor((Int32) color.Indexed.Value);
                else if (color.Theme != null)
                {
                    retVal = color.Tint != null ? XLColor.FromTheme((XLThemeColor) color.Theme.Value, color.Tint.Value) : XLColor.FromTheme((XLThemeColor) color.Theme.Value);
                }
            }
            return retVal ?? new XLColor();
        }

        private void ApplyStyle(IXLStylized xlStylized, Int32 styleIndex, Stylesheet s, Fills fills, Borders borders,
                                Fonts fonts, NumberingFormats numberingFormats)
        {
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