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
        public void Load(String file)
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

                //return items[int.Parse(headCell.CellValue.Text)].InnerText;

                var sheets = dSpreadsheet.WorkbookPart.Workbook.Sheets;
                
                // For each sheet, display the sheet information.
                foreach (var sheet in sheets)
                {
                    var dSheet = ((Sheet)sheet);
                    WorksheetPart worksheetPart = (WorksheetPart)dSpreadsheet.WorkbookPart.GetPartById(dSheet.Id);


                    var sheetName = dSheet.Name;


                    var ws = Worksheets.Add(sheetName);
                    foreach (var cell in worksheetPart.Worksheet.Descendants<Cell>())
                    {
                        var dCell = (Cell)cell;
                        if (dCell.DataType != null)
                        {
                            var xlCell = ws.Cell(dCell.CellReference);
                            Int32 styleIndex = dCell.StyleIndex != null ? Int32.Parse(dCell.StyleIndex.InnerText) : -1;
                            if (styleIndex >= 0)
                            {
                                styleIndex = Int32.Parse(dCell.StyleIndex.InnerText);
                                var fillId = ((CellFormat)((CellFormats)s.CellFormats).ElementAt(styleIndex)).FillId.Value;
                                var fill = (Fill)fills.ElementAt(Int32.Parse(fillId.ToString()));
                                xlCell.Style.Fill.PatternType = fillPatternValues.Single(p => p.Value == fill.PatternFill.PatternType).Key;
                                xlCell.Style.Fill.PatternColor = System.Drawing.ColorTranslator.FromHtml("#" + fill.PatternFill.ForegroundColor.Rgb.Value);
                                xlCell.Style.Fill.PatternBackgroundColor = System.Drawing.ColorTranslator.FromHtml("#" + fill.PatternFill.BackgroundColor.Rgb.Value);

                                var alignment = (Alignment)((CellFormat)((CellFormats)s.CellFormats).ElementAt(styleIndex)).Alignment;
                                xlCell.Style.Alignment.Horizontal = alignmentHorizontalValues.Single(a => a.Value == alignment.Horizontal).Key;
                                xlCell.Style.Alignment.Indent = Int32.Parse(alignment.Indent.ToString());
                                xlCell.Style.Alignment.JustifyLastLine = alignment.JustifyLastLine;
                                xlCell.Style.Alignment.ReadingOrder = (XLAlignmentReadingOrderValues)Int32.Parse(alignment.ReadingOrder.ToString());
                                xlCell.Style.Alignment.RelativeIndent = alignment.RelativeIndent;
                                xlCell.Style.Alignment.ShrinkToFit = alignment.ShrinkToFit;
                                xlCell.Style.Alignment.TextRotation = Int32.Parse(alignment.TextRotation.ToString());
                                xlCell.Style.Alignment.Vertical = alignmentVerticalValues.Single(a => a.Value == alignment.Vertical).Key;
                                xlCell.Style.Alignment.WrapText = alignment.WrapText;
                            }

                            if (dCell.DataType == CellValues.SharedString)
                            {
                                xlCell.DataType = XLCellValues.Text;
                                xlCell.Value = sharedStrings[Int32.Parse(dCell.CellValue.Text)].InnerText;
                            }
                            else if (dCell.DataType == CellValues.Date)
                            {
                                xlCell.DataType = XLCellValues.DateTime;
                                xlCell.Value = DateTime.FromOADate(Double.Parse(dCell.CellValue.Text)).ToString();
                            }
                            else if (dCell.DataType == CellValues.Boolean)
                            {
                                xlCell.DataType = XLCellValues.Boolean;
                                xlCell.Value = (dCell.CellValue.Text == "1").ToString();
                            }
                            else if (dCell.DataType == CellValues.Number)
                            {
                                xlCell.DataType = XLCellValues.Number;
                                xlCell.Value = dCell.CellValue.Text;
                                if (styleIndex >= 0)
                                {
                                    var numberFormatId = ((CellFormat)((CellFormats)s.CellFormats).ElementAt(styleIndex)).NumberFormatId;
                                    var numberFormatList = numberingFormats.Where(nf => ((NumberingFormat)nf).NumberFormatId.Value == numberFormatId);
                                    var formatCode = String.Empty;
                                    if (numberFormatList.Count() > 0)
                                    {
                                        NumberingFormat numberingFormat = (NumberingFormat)numberFormatList.First();
                                        formatCode = numberingFormat.FormatCode.Value;
                                    }
                                    if (formatCode.Length > 0)
                                        xlCell.Style.NumberFormat.Format = formatCode;
                                    else
                                        xlCell.Style.NumberFormat.NumberFormatId = Int32.Parse(numberFormatId);
                                }
                            }
                        }
                        //else if (dCell.CellValue !=null)
                        //{
                        //     var styleIndex = Int32.Parse(dCell.StyleIndex.InnerText);
                        //     var numberFormatId = ((CellFormat)((CellFormats)s.CellFormats).ElementAt(styleIndex)).NumberFormatId; //. [styleIndex].NumberFormatId;
                        //    ws.Cell(dCell.CellReference).Value = dCell.CellValue.Text;
                        //    ws.Cell(dCell.CellReference).Style.NumberFormat.NumberFormatId = Int32.Parse(numberFormatId);
                        //}
                    }
                }
            }
        }

    }
}