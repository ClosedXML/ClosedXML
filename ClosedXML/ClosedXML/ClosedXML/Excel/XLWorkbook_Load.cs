using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using C = DocumentFormat.OpenXml.Drawing.Charts;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;



namespace ClosedXML.Excel
{
    public partial class XLWorkbook
    {
        private void Load(String file)
        {

            LoadSheets(file);
        }

        private void LoadSheets(String fileName)
        {
            // Open file as read-only.
            using (SpreadsheetDocument dSpreadsheet = SpreadsheetDocument.Open(fileName, false))
            {
                SharedStringItem[] sharedStrings = null;
                if (dSpreadsheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                {
                    SharedStringTablePart shareStringPart = dSpreadsheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    sharedStrings = shareStringPart.SharedStringTable.Elements<SharedStringItem>().ToArray();
                }

                var workbookStylesPart = (WorkbookStylesPart)dSpreadsheet.WorkbookPart.WorkbookStylesPart;
                var s = (Stylesheet)workbookStylesPart.Stylesheet;
                var numberingFormats = (NumberingFormats)s.NumberingFormats;
                Fills fills = (Fills)s.Fills;
                Borders borders = (Borders)s.Borders;
                Fonts fonts = (Fonts)s.Fonts;
               

                //return items[int.Parse(headCell.CellValue.Text)].InnerText;

                var sheets = dSpreadsheet.WorkbookPart.Workbook.Sheets;
                
                // For each sheet, display the sheet information.
                foreach (var sheet in sheets)
                {
                    var dSheet = ((Sheet)sheet);
                    WorksheetPart worksheetPart = (WorksheetPart)dSpreadsheet.WorkbookPart.GetPartById(dSheet.Id);
                    
                    var sheetName = dSheet.Name;

                    var ws = (XLWorksheet)Worksheets.Add(sheetName);

                    var sheetFormatProperties = (SheetFormatProperties)worksheetPart.Worksheet.Descendants<SheetFormatProperties>().First();
                    ws.RowHeight = sheetFormatProperties.DefaultRowHeight;
                    ws.ColumnWidth = sheetFormatProperties.DefaultColumnWidth;

                    foreach (var mCell in worksheetPart.Worksheet.Descendants<MergeCell>())
                    {
                        var mergeCell = (MergeCell)mCell;
                        ws.Range(mergeCell.Reference).Merge();
                    }


                    var wsDefaultColumn = worksheetPart.Worksheet.Descendants<Column>().Where(
                        c => c.Max == XLWorksheet.MaxNumberOfColumns).Single();

                    if (wsDefaultColumn.Width != null) ws.ColumnWidth = wsDefaultColumn.Width;

                    Int32 styleIndexDefault = wsDefaultColumn.Style != null ? Int32.Parse(wsDefaultColumn.Style.InnerText) : -1;
                    if (styleIndexDefault >= 0)
                    {
                        ApplyStyle(ws, styleIndexDefault, s, fills, borders, fonts, numberingFormats);
                    }

                    foreach (var col in worksheetPart.Worksheet.Descendants<Column>())
                    {
                        IXLStylized toApply;
                        if (col.Max != XLWorksheet.MaxNumberOfColumns)
                        {
                            toApply = ws.Columns(col.Min, col.Max);
                            if (col.Width != null) 
                                ((XLColumns)toApply).Width = col.Width;
                            else
                                ((XLColumns)toApply).Width = ws.ColumnWidth;

                            Int32 styleIndex = col.Style != null ? Int32.Parse(col.Style.InnerText) : -1;
                            if (styleIndex > 0)
                            {
                                ApplyStyle(toApply, styleIndex, s, fills, borders, fonts, numberingFormats);
                            }
                            else
                            {
                                toApply.Style = DefaultStyle;
                            }
                        }
                    }

                    foreach (var row in worksheetPart.Worksheet.Descendants<Row>().Where(r=>r.CustomFormat != null && r.CustomFormat).Select(r=>r))
                    {
                        //var dRow = (Column)col;
                        var xlRow = ws.Row(Int32.Parse(row.RowIndex.ToString()));
                        if (row.Height != null)
                            xlRow.Height = row.Height;
                        else
                            xlRow.Height = ws.RowHeight;

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


                    foreach (var cell in worksheetPart.Worksheet.Descendants<Cell>())
                    {
                        var dCell = (Cell)cell;
                        Int32 styleIndex = dCell.StyleIndex != null ? Int32.Parse(dCell.StyleIndex.InnerText) : -1;
                        var xlCell = ws.Cell(dCell.CellReference);
                        if (styleIndex > 0)
                        {
                            styleIndex = Int32.Parse(dCell.StyleIndex.InnerText);
                            ApplyStyle(xlCell, styleIndex, s, fills, borders, fonts, numberingFormats);
                        }
                        else
                        {
                            xlCell.Style = DefaultStyle;
                        }

                        if (dCell.DataType != null)
                        {
                            if (dCell.DataType == CellValues.SharedString)
                            {
                                xlCell.DataType = XLCellValues.Text;
                                xlCell.Value = sharedStrings[Int32.Parse(dCell.CellValue.Text)].InnerText;
                            }
                            else if (dCell.DataType == CellValues.Date)
                            {
                                xlCell.DataType = XLCellValues.DateTime;
                                xlCell.Value = DateTime.FromOADate(Double.Parse(dCell.CellValue.Text));
                            }
                            else if (dCell.DataType == CellValues.Boolean)
                            {
                                xlCell.DataType = XLCellValues.Boolean;
                                xlCell.Value = (dCell.CellValue.Text == "1");
                            }
                            else if (dCell.DataType == CellValues.Number)
                            {
                                xlCell.DataType = XLCellValues.Number;
                                xlCell.Value = dCell.CellValue.Text;
                            }
                        }
                        else if (dCell.CellValue != null)
                        {
                            //var styleIndex = Int32.Parse(dCell.StyleIndex.InnerText);
                            var numberFormatId = ((CellFormat)((CellFormats)s.CellFormats).ElementAt(styleIndex)).NumberFormatId; //. [styleIndex].NumberFormatId;
                            ws.Cell(dCell.CellReference).Value = dCell.CellValue.Text;
                            ws.Cell(dCell.CellReference).Style.NumberFormat.NumberFormatId = Int32.Parse(numberFormatId);
                        }
                    }

                    var printOptions = (PrintOptions)worksheetPart.Worksheet.Descendants<PrintOptions>().First();
                    ws.PageSetup.ShowGridlines = printOptions.GridLines;
                    ws.PageSetup.CenterHorizontally = printOptions.HorizontalCentered;
                    ws.PageSetup.CenterVertically = printOptions.VerticalCentered;
                    ws.PageSetup.ShowRowAndColumnHeadings = printOptions.Headings;

                    var pageMargins = (PageMargins)worksheetPart.Worksheet.Descendants<PageMargins>().First();
                    ws.PageSetup.Margins.Bottom = pageMargins.Bottom;
                    ws.PageSetup.Margins.Footer = pageMargins.Footer;
                    ws.PageSetup.Margins.Header = pageMargins.Header;
                    ws.PageSetup.Margins.Left = pageMargins.Left;
                    ws.PageSetup.Margins.Right = pageMargins.Right;
                    ws.PageSetup.Margins.Top = pageMargins.Top;

                    var pageSetup = (PageSetup)worksheetPart.Worksheet.Descendants<PageSetup>().First();
                    ws.PageSetup.PaperSize = (XLPaperSize)Int32.Parse(pageSetup.PaperSize.InnerText);
                    if (pageSetup.Scale != null)
                    {
                        ws.PageSetup.Scale = Int32.Parse(pageSetup.Scale.InnerText);
                    }
                    else
                    {
                        ws.PageSetup.FitToPages(Int32.Parse(pageSetup.FitToWidth.InnerText), Int32.Parse(pageSetup.FitToHeight.InnerText));
                    }
                    ws.PageSetup.PageOrder = pageOrderValues.Single(p => p.Value == pageSetup.PageOrder).Key;
                    ws.PageSetup.PageOrientation = pageOrientationValues.Single(p => p.Value == pageSetup.Orientation).Key;
                    ws.PageSetup.BlackAndWhite = pageSetup.BlackAndWhite;
                    ws.PageSetup.DraftQuality = pageSetup.Draft;
                    ws.PageSetup.ShowComments = showCommentsValues.Single(sc => sc.Value == pageSetup.CellComments).Key;
                    ws.PageSetup.PrintErrorValue = printErrorValues.Single(p => p.Value == pageSetup.Errors).Key;
                    if (pageSetup.HorizontalDpi != null) ws.PageSetup.HorizontalDpi = Int32.Parse(pageSetup.HorizontalDpi.InnerText);
                    if (pageSetup.VerticalDpi != null) ws.PageSetup.VerticalDpi = Int32.Parse(pageSetup.VerticalDpi.InnerText);
                    if (pageSetup.FirstPageNumber != null) ws.PageSetup.FirstPageNumber = Int32.Parse(pageSetup.FirstPageNumber.InnerText);

                    var headerFooter = (HeaderFooter)worksheetPart.Worksheet.Descendants<HeaderFooter>().First();
                    ws.PageSetup.AlignHFWithMargins = headerFooter.AlignWithMargins;
                    // Footers
                    var xlFooter = (XLHeaderFooter)ws.PageSetup.Footer;
                    var evenFooter = (EvenFooter)headerFooter.EvenFooter;
                    xlFooter.SetInnerText(XLHFOccurrence.EvenPages, evenFooter.Text);
                    var oddFooter = (OddFooter)headerFooter.OddFooter;
                    xlFooter.SetInnerText(XLHFOccurrence.OddPages, oddFooter.Text);
                    var firstFooter = (FirstFooter)headerFooter.FirstFooter;
                    xlFooter.SetInnerText(XLHFOccurrence.FirstPage, firstFooter.Text);
                    // Headers
                    var xlHeader = (XLHeaderFooter)ws.PageSetup.Header;
                    var evenHeader = (EvenHeader)headerFooter.EvenHeader;
                    xlHeader.SetInnerText(XLHFOccurrence.EvenPages, evenHeader.Text);
                    var oddHeader = (OddHeader)headerFooter.OddHeader;
                    xlHeader.SetInnerText(XLHFOccurrence.OddPages, oddHeader.Text);
                    var firstHeader = (FirstHeader)headerFooter.FirstHeader;
                    xlHeader.SetInnerText(XLHFOccurrence.FirstPage, firstHeader.Text);

                    var rowBreaksList = worksheetPart.Worksheet.Descendants<RowBreaks>();
                    if (rowBreaksList.Count() > 0)
                    {
                        var rowBreaks = (RowBreaks)rowBreaksList.First();
                        foreach (var rowBreak in rowBreaks.Descendants<Break>())
                        {
                            ws.PageSetup.RowBreaks.Add(Int32.Parse(rowBreak.Id.InnerText));
                        }
                    }

                    var columnBreaksList = worksheetPart.Worksheet.Descendants<ColumnBreaks>();
                    if (columnBreaksList.Count() > 0)
                    {
                        var columnBreaks = (ColumnBreaks)columnBreaksList.First();
                        foreach (var columnBreak in columnBreaks.Descendants<Break>())
                        {
                            ws.PageSetup.ColumnBreaks.Add(Int32.Parse(columnBreak.Id.InnerText));
                        }
                    }
                }

                var workbook = (Workbook)dSpreadsheet.WorkbookPart.Workbook;
                foreach (var definedName in workbook.Descendants<DefinedName>())
                {
                    if (definedName.Name == "_xlnm.Print_Area")
                    {
                        foreach (var area in definedName.Text.Split(','))
                        {
                            var sections = area.Split('!');
                            var sheetName = sections[0].Replace("\'", "");
                            var sheetArea = sections[1];
                            Worksheets.GetWorksheet(sheetName).PageSetup.PrintAreas.Add(sheetArea);
                        }
                    }
                    else if (definedName.Name == "_xlnm.Print_Titles")
                    {
                        var areas = definedName.Text.Split(',');

                        var colSections = areas[0].Split('!');
                        var sheetNameCol = colSections[0].Replace("\'", "");
                        var sheetAreaCol = colSections[1];
                        Worksheets.GetWorksheet(sheetNameCol).PageSetup.SetColumnsToRepeatAtLeft(sheetAreaCol);

                        var rowSections = areas[1].Split('!');
                        var sheetNameRow = rowSections[0].Replace("\'", "");
                        var sheetAreaRow = rowSections[1];
                        Worksheets.GetWorksheet(sheetNameRow).PageSetup.SetRowsToRepeatAtTop(sheetAreaRow);
                    }
                    //ws.PageSetup.PrintAreas.
                }
            }
        }

        private void ApplyStyle(IXLStylized xlStylized, Int32 styleIndex, Stylesheet s, Fills fills, Borders borders, Fonts fonts, NumberingFormats numberingFormats )
        {
            var fillId = ((CellFormat)((CellFormats)s.CellFormats).ElementAt(styleIndex)).FillId.Value;
            if (fillId > 0)
            {
                var fill = (Fill)fills.ElementAt(Int32.Parse(fillId.ToString()));
                xlStylized.Style.Fill.PatternType = fillPatternValues.Single(p => p.Value == fill.PatternFill.PatternType).Key;
                xlStylized.Style.Fill.PatternColor = System.Drawing.ColorTranslator.FromHtml("#" + fill.PatternFill.ForegroundColor.Rgb.Value);
                xlStylized.Style.Fill.PatternBackgroundColor = System.Drawing.ColorTranslator.FromHtml("#" + fill.PatternFill.BackgroundColor.Rgb.Value);
            }

            var alignment = (Alignment)((CellFormat)((CellFormats)s.CellFormats).ElementAt(styleIndex)).Alignment;
            xlStylized.Style.Alignment.Horizontal = alignmentHorizontalValues.Single(a => a.Value == alignment.Horizontal).Key;
            xlStylized.Style.Alignment.Indent = Int32.Parse(alignment.Indent.ToString());
            xlStylized.Style.Alignment.JustifyLastLine = alignment.JustifyLastLine;
            xlStylized.Style.Alignment.ReadingOrder = (XLAlignmentReadingOrderValues)Int32.Parse(alignment.ReadingOrder.ToString());
            xlStylized.Style.Alignment.RelativeIndent = alignment.RelativeIndent;
            xlStylized.Style.Alignment.ShrinkToFit = alignment.ShrinkToFit;
            xlStylized.Style.Alignment.TextRotation = Int32.Parse(alignment.TextRotation.ToString());
            xlStylized.Style.Alignment.Vertical = alignmentVerticalValues.Single(a => a.Value == alignment.Vertical).Key;
            xlStylized.Style.Alignment.WrapText = alignment.WrapText;

            var borderId = ((CellFormat)((CellFormats)s.CellFormats).ElementAt(styleIndex)).BorderId.Value;
            var border = (Border)borders.ElementAt(Int32.Parse(borderId.ToString()));
            var bottomBorder = (BottomBorder)border.BottomBorder;
            xlStylized.Style.Border.BottomBorder = borderStyleValues.Single(b => b.Value == bottomBorder.Style.Value).Key;
            xlStylized.Style.Border.BottomBorderColor = System.Drawing.ColorTranslator.FromHtml("#" + ((Color)bottomBorder.Color).Rgb.Value);
            var topBorder = (TopBorder)border.TopBorder;
            xlStylized.Style.Border.TopBorder = borderStyleValues.Single(b => b.Value == topBorder.Style.Value).Key;
            xlStylized.Style.Border.TopBorderColor = System.Drawing.ColorTranslator.FromHtml("#" + ((Color)topBorder.Color).Rgb.Value);
            var leftBorder = (LeftBorder)border.LeftBorder;
            xlStylized.Style.Border.LeftBorder = borderStyleValues.Single(b => b.Value == leftBorder.Style.Value).Key;
            xlStylized.Style.Border.LeftBorderColor = System.Drawing.ColorTranslator.FromHtml("#" + ((Color)leftBorder.Color).Rgb.Value);
            var rightBorder = (RightBorder)border.RightBorder;
            xlStylized.Style.Border.RightBorder = borderStyleValues.Single(b => b.Value == rightBorder.Style.Value).Key;
            xlStylized.Style.Border.RightBorderColor = System.Drawing.ColorTranslator.FromHtml("#" + ((Color)rightBorder.Color).Rgb.Value);
            var diagonalBorder = (DiagonalBorder)border.DiagonalBorder;
            xlStylized.Style.Border.DiagonalBorder = borderStyleValues.Single(b => b.Value == diagonalBorder.Style.Value).Key;
            xlStylized.Style.Border.DiagonalBorderColor = System.Drawing.ColorTranslator.FromHtml("#" + ((Color)diagonalBorder.Color).Rgb.Value);
            xlStylized.Style.Border.DiagonalDown = border.DiagonalDown;
            xlStylized.Style.Border.DiagonalUp = border.DiagonalUp;

            var fontId = ((CellFormat)((CellFormats)s.CellFormats).ElementAt(styleIndex)).FontId;
            var font = (Font)fonts.ElementAt(Int32.Parse(fontId.ToString()));
            xlStylized.Style.Font.Bold = (font.Bold != null);
            xlStylized.Style.Font.FontColor = System.Drawing.ColorTranslator.FromHtml("#" + ((Color)font.Color).Rgb.Value);
            xlStylized.Style.Font.FontFamilyNumbering = (XLFontFamilyNumberingValues)Int32.Parse(((FontFamilyNumbering)font.FontFamilyNumbering).Val.ToString());
            xlStylized.Style.Font.FontName = ((FontName)font.FontName).Val;
            xlStylized.Style.Font.FontSize = ((FontSize)font.FontSize).Val;
            xlStylized.Style.Font.Italic = (font.Italic != null);
            xlStylized.Style.Font.Shadow = (font.Shadow != null);
            xlStylized.Style.Font.Strikethrough = (font.Strike != null);
            xlStylized.Style.Font.Underline = font.Underline == null || ((Underline)font.Underline).Val == null ? XLWorkbook.DefaultStyle.Font.Underline : underlineValuesList.Single(u => u.Value == ((Underline)font.Underline).Val).Key;
            xlStylized.Style.Font.VerticalAlignment = fontVerticalTextAlignmentValues.Single(f => f.Value == ((VerticalTextAlignment)font.VerticalTextAlignment).Val).Key;

            var numberFormatId = ((CellFormat)((CellFormats)s.CellFormats).ElementAt(styleIndex)).NumberFormatId;
            var numberFormatList = numberingFormats.Where(nf => ((NumberingFormat)nf).NumberFormatId.Value == numberFormatId);
            var formatCode = String.Empty;
            if (numberFormatList.Count() > 0)
            {
                NumberingFormat numberingFormat = (NumberingFormat)numberFormatList.First();
                formatCode = numberingFormat.FormatCode.Value;
            }
            if (formatCode.Length > 0)
                xlStylized.Style.NumberFormat.Format = formatCode;
            else
                xlStylized.Style.NumberFormat.NumberFormatId = Int32.Parse(numberFormatId);
        }

    }
}