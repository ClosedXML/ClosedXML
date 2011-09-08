using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.VariantTypes;
using Vml = DocumentFormat.OpenXml.Vml;
using BackgroundColor = DocumentFormat.OpenXml.Spreadsheet.BackgroundColor;
using BottomBorder = DocumentFormat.OpenXml.Spreadsheet.BottomBorder;
using Break = DocumentFormat.OpenXml.Spreadsheet.Break;
using Fill = DocumentFormat.OpenXml.Spreadsheet.Fill;
using FontScheme = DocumentFormat.OpenXml.Drawing.FontScheme;
using Fonts = DocumentFormat.OpenXml.Spreadsheet.Fonts;
using ForegroundColor = DocumentFormat.OpenXml.Spreadsheet.ForegroundColor;
using GradientFill = DocumentFormat.OpenXml.Drawing.GradientFill;
using GradientStop = DocumentFormat.OpenXml.Drawing.GradientStop;
using Hyperlink = DocumentFormat.OpenXml.Spreadsheet.Hyperlink;
using LeftBorder = DocumentFormat.OpenXml.Spreadsheet.LeftBorder;
using Outline = DocumentFormat.OpenXml.Drawing.Outline;
using Path = System.IO.Path;
using PatternFill = DocumentFormat.OpenXml.Spreadsheet.PatternFill;
using Properties = DocumentFormat.OpenXml.ExtendedProperties.Properties;
using RightBorder = DocumentFormat.OpenXml.Spreadsheet.RightBorder;
using Table = DocumentFormat.OpenXml.Spreadsheet.Table;
using Text = DocumentFormat.OpenXml.Spreadsheet.Text;
using TopBorder = DocumentFormat.OpenXml.Spreadsheet.TopBorder;
using Underline = DocumentFormat.OpenXml.Spreadsheet.Underline;


namespace ClosedXML.Excel
{
    public partial class XLWorkbook
    {
        private const Double ColumnWidthOffset = 0.710625;

        //private Dictionary<String, UInt32> sharedStrings;
        //private Dictionary<IXLStyle, StyleInfo> context.SharedStyles;

        private static readonly EnumValue<CellValues> CvSharedString = new EnumValue<CellValues>(CellValues.SharedString);
        private static readonly EnumValue<CellValues> CvInlineString = new EnumValue<CellValues>(CellValues.InlineString);
        private static readonly EnumValue<CellValues> CvNumber = new EnumValue<CellValues>(CellValues.Number);
        private static readonly EnumValue<CellValues> CvDate = new EnumValue<CellValues>(CellValues.Date);
        private static readonly EnumValue<CellValues> CvBoolean = new EnumValue<CellValues>(CellValues.Boolean);

        private static EnumValue<CellValues> GetCellValue(XLCell xlCell)
        {
            switch (xlCell.DataType)
            {
                case XLCellValues.Text:
                    {
                        return xlCell.ShareString ? CvSharedString : CvInlineString;
                    }
                case XLCellValues.Number:
                    return CvNumber;
                case XLCellValues.DateTime:
                    return CvDate;
                case XLCellValues.Boolean:
                    return CvBoolean;
                case XLCellValues.TimeSpan:
                    return CvNumber;
                default:
                    throw new NotImplementedException();
            }
        }

        private void CreatePackage(String filePath)
        {
            PathHelper.CreateDirectory(Path.GetDirectoryName(filePath));
            var package = File.Exists(filePath)
                              ? SpreadsheetDocument.Open(filePath, true)
                              : SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);

            using (package)
            {
                CreateParts(package);
                //package.Close();
            }
        }

        private void CreatePackage(Stream stream, Boolean newStream)
        {
            var package = newStream
                              ? SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook)
                              : SpreadsheetDocument.Open(stream, true);

            using (package)
            {
                CreateParts(package);
                //package.Close();
            }
        }

        // Adds child parts and generates content of the specified part.
        private void CreateParts(SpreadsheetDocument document)
        {
            var context = new SaveContext();

            var workbookPart = document.WorkbookPart ?? document.AddWorkbookPart();

            var worksheets = WorksheetsInternal;
            var partsToRemove = workbookPart.Parts.Where(s => worksheets.Deleted.Contains(s.RelationshipId)).ToList();
            partsToRemove.ForEach(s => workbookPart.DeletePart(s.OpenXmlPart));
            context.RelIdGenerator.AddValues(workbookPart.Parts.Select(p => p.RelationshipId).ToList(), RelType.Workbook);

            var extendedFilePropertiesPart = document.ExtendedFilePropertiesPart ??
                                             document.AddNewPart<ExtendedFilePropertiesPart>(
                                                 context.RelIdGenerator.GetNext(RelType.Workbook));

            GenerateExtendedFilePropertiesPartContent(extendedFilePropertiesPart);

            GenerateWorkbookPartContent(workbookPart, context);

            var sharedStringTablePart = workbookPart.SharedStringTablePart ??
                                        workbookPart.AddNewPart<SharedStringTablePart>(
                                            context.RelIdGenerator.GetNext(RelType.Workbook));

            GenerateSharedStringTablePartContent(sharedStringTablePart, context);

            var workbookStylesPart = workbookPart.WorkbookStylesPart ??
                                     workbookPart.AddNewPart<WorkbookStylesPart>(
                                         context.RelIdGenerator.GetNext(RelType.Workbook));

            GenerateWorkbookStylesPartContent(workbookStylesPart, context);

            foreach (XLWorksheet worksheet in WorksheetsInternal.Cast<XLWorksheet>().OrderBy(w => w.Position))
            {
                WorksheetPart worksheetPart;
                string wsRelId = worksheet.RelId;
                if (workbookPart.Parts.Any(p => p.RelationshipId == wsRelId))
                {
                    worksheetPart = (WorksheetPart)workbookPart.GetPartById(wsRelId);
                    var wsPartsToRemove = worksheetPart.TableDefinitionParts.ToList();
                    wsPartsToRemove.ForEach(tdp => worksheetPart.DeletePart(tdp));
                }
                else
                    worksheetPart = workbookPart.AddNewPart<WorksheetPart>(wsRelId);

                context.RelIdGenerator.AddValues(worksheetPart.Parts.Select(p => p.RelationshipId).ToList(), RelType.Worksheet);

                // delete comment related parts (todo: review)
                //worksheetPart.DeletePart(worksheetPart.WorksheetCommentsPart);
                //worksheetPart.DeleteParts<VmlDrawingPart>(worksheetPart.GetPartsOfType<VmlDrawingPart>());

                //if (worksheet.Internals.CellsCollection.GetCells(c => c.HasComment).Any())
                //{
                //    WorksheetCommentsPart worksheetCommentsPart =
                //        worksheetPart.AddNewPart<WorksheetCommentsPart>(context.RelIdGenerator.GetNext(RelType.Worksheet));
                //    GenerateWorksheetCommentsPartContent(worksheetCommentsPart, worksheet);

                //    worksheet.LegacyDrawingId = context.RelIdGenerator.GetNext(RelType.Worksheet);
                //    VmlDrawingPart vmlDrawingPart = worksheetPart.AddNewPart<VmlDrawingPart>(worksheet.LegacyDrawingId);
                //    GenerateVmlDrawingPartContent(vmlDrawingPart, worksheet, context);
                //}

                GenerateWorksheetPartContent(worksheetPart, worksheet, context);

                //if (worksheet.PivotTables.Any())
                //{
                //    GeneratePivotTables(workbookPart, worksheetPart, worksheet, context);
                //}



                //DrawingsPart drawingsPart = worksheetPart.AddNewPart<DrawingsPart>("rId1");
                //GenerateDrawingsPartContent(drawingsPart, worksheet);

                //foreach (var chart in worksheet.Charts)
                //{
                //    ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>("rId1");
                //    GenerateChartPartContent(chartPart, (XLChart)chart);
                //}
            }

            GenerateCalculationChainPartContent(workbookPart, context);

            if (workbookPart.ThemePart == null)
            {
                var themePart = workbookPart.AddNewPart<ThemePart>(context.RelIdGenerator.GetNext(RelType.Workbook));
                GenerateThemePartContent(themePart);
            }

            if (CustomProperties.Any())
            {
                document.GetPartsOfType<CustomFilePropertiesPart>().ToList().ForEach(p => document.DeletePart(p));
                var customFilePropertiesPart =
                    document.AddNewPart<CustomFilePropertiesPart>(context.RelIdGenerator.GetNext(RelType.Workbook));

                GenerateCustomFilePropertiesPartContent(customFilePropertiesPart);
            }
            SetPackageProperties(document);
        }

        private static void GenerateTables(XLWorksheet worksheet, WorksheetPart worksheetPart, SaveContext context)
        {
            worksheetPart.Worksheet.RemoveAllChildren<TablePart>();

            if (!worksheet.Tables.Any()) return;

            foreach (IXLTable table in worksheet.Tables)
            {
                string tableRelId = context.RelIdGenerator.GetNext(RelType.Workbook);
                var xlTable = (XLTable)table;
                xlTable.RelId = tableRelId;
                var tableDefinitionPart = worksheetPart.AddNewPart<TableDefinitionPart>(tableRelId);
                GenerateTableDefinitionPartContent(tableDefinitionPart, xlTable, context);
            }
        }

        private void GenerateExtendedFilePropertiesPartContent(ExtendedFilePropertiesPart extendedFilePropertiesPart)
        {
            if (extendedFilePropertiesPart.Properties == null)
                extendedFilePropertiesPart.Properties = new Properties();

            var properties = extendedFilePropertiesPart.Properties;
            if (
                !properties.NamespaceDeclarations.Contains(new KeyValuePair<string, string>("vt",
                                                                                            "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")))
            {
                properties.AddNamespaceDeclaration("vt",
                                                   "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            }

            if (properties.Application == null)
                properties.AppendChild(new Application {Text = "Microsoft Excel"});

            if (properties.DocumentSecurity == null)
                properties.AppendChild(new DocumentSecurity {Text = "0"});

            if (properties.ScaleCrop == null)
                properties.AppendChild(new ScaleCrop {Text = "false"});

            if (properties.HeadingPairs == null)
                properties.HeadingPairs = new HeadingPairs();

            if (properties.TitlesOfParts == null)
                properties.TitlesOfParts = new TitlesOfParts();

            properties.HeadingPairs.VTVector = new VTVector {BaseType = VectorBaseValues.Variant};

            properties.TitlesOfParts.VTVector = new VTVector {BaseType = VectorBaseValues.Lpstr};

            var vTVectorOne = properties.HeadingPairs.VTVector;

            var vTVectorTwo = properties.TitlesOfParts.VTVector;

            var modifiedWorksheets =
                ((IEnumerable<XLWorksheet>)WorksheetsInternal).Select(w => new {w.Name, Order = w.Position}).ToList();
            var modifiedNamedRanges = GetModifiedNamedRanges();
            int modifiedWorksheetsCount = modifiedWorksheets.Count;
            int modifiedNamedRangesCount = modifiedNamedRanges.Count;

            InsertOnVtVector(vTVectorOne, "Worksheets", 0, modifiedWorksheetsCount.ToString());
            InsertOnVtVector(vTVectorOne, "Named Ranges", 2, modifiedNamedRangesCount.ToString());

            vTVectorTwo.Size = (UInt32)(modifiedNamedRangesCount + modifiedWorksheetsCount);

            foreach (
                VTLPSTR vTlpstr3 in modifiedWorksheets.OrderBy(w => w.Order).Select(w => new VTLPSTR {Text = w.Name}))
                vTVectorTwo.AppendChild(vTlpstr3);

            foreach (VTLPSTR vTlpstr7 in modifiedNamedRanges.Select(nr => new VTLPSTR {Text = nr}))
                vTVectorTwo.AppendChild(vTlpstr7);

            if (Properties.Manager != null)
            {
                if (!StringExtensions.IsNullOrWhiteSpace(Properties.Manager))
                {
                    if (properties.Manager == null)
                        properties.Manager = new Manager();

                    properties.Manager.Text = Properties.Manager;
                }
                else
                    properties.Manager = null;
            }

            if (Properties.Company == null) return;

            if (!StringExtensions.IsNullOrWhiteSpace(Properties.Company))
            {
                if (properties.Company == null)
                    properties.Company = new Company();

                properties.Company.Text = Properties.Company;
            }
            else
                properties.Company = null;
        }

        private static void InsertOnVtVector(VTVector vTVector, String property, Int32 index, String text)
        {
            var m = from e1 in vTVector.Elements<Variant>()
                    where e1.Elements<VTLPSTR>().Any(e2 => e2.Text == property)
                    select e1;
            if (!m.Any())
            {
                if (vTVector.Size == null)
                    vTVector.Size = new UInt32Value(0U);

                vTVector.Size += 2U;
                var variant1 = new Variant();
                var vTlpstr1 = new VTLPSTR {Text = property};
                variant1.AppendChild(vTlpstr1);
                vTVector.InsertAt(variant1, index);

                var variant2 = new Variant();
                var vTInt321 = new VTInt32();
                variant2.AppendChild(vTInt321);
                vTVector.InsertAt(variant2, index + 1);
            }

            Int32 targetIndex = 0;
            foreach (Variant e in vTVector.Elements<Variant>())
            {
                if (e.Elements<VTLPSTR>().Any(e2 => e2.Text == property))
                {
                    vTVector.ElementAt(targetIndex + 1).GetFirstChild<VTInt32>().Text = text;
                    break;
                }
                targetIndex++;
            }
        }

        private List<string> GetModifiedNamedRanges()
        {
            var namedRanges = new List<String>();
            foreach (XLWorksheet w in WorksheetsInternal)
            {
                String wName = w.Name;
                namedRanges.AddRange(w.NamedRanges.Select(n => wName + "!" + n.Name));
                namedRanges.Add(w.Name + "!Print_Area");
                namedRanges.Add(w.Name + "!Print_Titles");
            }
            namedRanges.AddRange(NamedRanges.Select(n => n.Name));
            return namedRanges;
        }

        private void GenerateWorkbookPartContent(WorkbookPart workbookPart, SaveContext context)
        {
            if (workbookPart.Workbook == null)
                workbookPart.Workbook = new Workbook();

            var workbook = workbookPart.Workbook;
            if (
                !workbook.NamespaceDeclarations.Contains(new KeyValuePair<string, string>("r",
                                                                                          "http://schemas.openxmlformats.org/officeDocument/2006/relationships")))
            {
                workbook.AddNamespaceDeclaration("r",
                                                 "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            }

            #region WorkbookProperties

            if (workbook.WorkbookProperties == null)
                workbook.WorkbookProperties = new WorkbookProperties();

            if (workbook.WorkbookProperties.CodeName == null)
                workbook.WorkbookProperties.CodeName = "ThisWorkbook";

            if (workbook.WorkbookProperties.DefaultThemeVersion == null)
                workbook.WorkbookProperties.DefaultThemeVersion = 124226U;

            #endregion

            if (workbook.BookViews == null)
                workbook.BookViews = new BookViews();

            if (workbook.Sheets == null)
                workbook.Sheets = new Sheets();

            var worksheets = WorksheetsInternal;
            workbook.Sheets.Elements<Sheet>().Where(s => worksheets.Deleted.Contains(s.Id)).ToList().ForEach(
                s => s.Remove());

            foreach (Sheet sheet in workbook.Sheets.Elements<Sheet>())
            {
                int sheetId = (Int32)sheet.SheetId.Value;

                if (!WorksheetsInternal.Any<XLWorksheet>(w => w.SheetId == sheetId)) continue;

                var wks =
                    WorksheetsInternal.Where<XLWorksheet>(w => w.SheetId == sheetId).Single();
                wks.RelId = sheet.Id;
                sheet.Name = wks.Name;
            }

            foreach (
                XLWorksheet xlSheet in
                    WorksheetsInternal.Cast<XLWorksheet>().Where(w => w.SheetId == 0).OrderBy(w => w.Position))
            {
                String rId = context.RelIdGenerator.GetNext(RelType.Workbook);
                //Int32 rIdSub = Int32.Parse(rId.Substring(3));
                while (WorksheetsInternal.Cast<XLWorksheet>().Any(w => w.SheetId == Int32.Parse(rId.Substring(3))))
                    rId = context.RelIdGenerator.GetNext(RelType.Workbook);

                xlSheet.SheetId = Int32.Parse(rId.Substring(3));
                xlSheet.RelId = rId;
                var newSheet = new Sheet
                                   {
                                       Name = xlSheet.Name,
                                       Id = rId,
                                       SheetId = (UInt32)xlSheet.SheetId
                                   };

                if (xlSheet.Visibility != XLWorksheetVisibility.Visible)
                    newSheet.State = xlSheet.Visibility.ToOpenXml();

                workbook.Sheets.AppendChild(newSheet);
            }

            var sheetElements = from sheet in workbook.Sheets.Elements<Sheet>()
                                join worksheet in ((IEnumerable<XLWorksheet>)WorksheetsInternal) on sheet.Id.Value equals worksheet.RelId
                                orderby worksheet.Position
                                select sheet;

            UInt32 firstSheetVisible = 0;
            UInt32 activeTab = (from us in _unsupportedSheets where us.Value.IsActive select (UInt32)us.Key - 1).FirstOrDefault();
            Boolean foundVisible = false;
            Int32 position = 0;
            foreach (Sheet sheet in sheetElements)
            {
                position++;
                if (_unsupportedSheets.ContainsKey(position))
                {
                    Sheet unsupportedSheet =
                        workbook.Sheets.Elements<Sheet>().Where(s => s.SheetId == _unsupportedSheets[position].SheetId).First();
                    workbook.Sheets.RemoveChild(unsupportedSheet);
                    workbook.Sheets.AppendChild(unsupportedSheet);
                    _unsupportedSheets.Remove(position);
                }
                
                    workbook.Sheets.RemoveChild(sheet);
                    workbook.Sheets.AppendChild(sheet);

                    if (foundVisible) continue;

                    if (sheet.State == null || sheet.State == SheetStateValues.Visible)
                        foundVisible = true;
                    else
                        firstSheetVisible++;
                
            }
            foreach (Sheet unsupportedSheet in _unsupportedSheets.Values.Select(us => workbook.Sheets.Elements<Sheet>().Where(s => s.SheetId == us.SheetId).First()))
            {
                workbook.Sheets.RemoveChild(unsupportedSheet);
                workbook.Sheets.AppendChild(unsupportedSheet);
            }

            var workbookView = workbook.BookViews.Elements<WorkbookView>().FirstOrDefault();

            if (activeTab == 0)
            {
                activeTab = firstSheetVisible;
                foreach (XLWorksheet ws in worksheets)
                {
                    if (!ws.TabActive) continue;

                    activeTab = (UInt32)(ws.Position - 1);
                    break;
                }
            }

            if (workbookView == null)
            {
                workbookView = new WorkbookView {ActiveTab = activeTab, FirstSheet = firstSheetVisible};
                workbook.BookViews.AppendChild(workbookView);
            }
            else
            {
                workbookView.ActiveTab = activeTab;
                workbookView.FirstSheet = firstSheetVisible;
            }

            var definedNames = new DefinedNames();
            foreach (XLWorksheet worksheet in WorksheetsInternal)
            {
                uint wsSheetId = (UInt32)worksheet.SheetId;
                UInt32 sheetId = 0;
                foreach (Sheet s in workbook.Sheets.Elements<Sheet>().TakeWhile(s => s.SheetId != wsSheetId))
                {
                    sheetId++;
                }

                if (worksheet.PageSetup.PrintAreas.Any())
                {
                    var definedName = new DefinedName {Name = "_xlnm.Print_Area", LocalSheetId = sheetId};
                    String worksheetName = worksheet.Name;
                    string definedNameText = worksheet.PageSetup.PrintAreas.Aggregate(String.Empty,
                                                                                      (current, printArea) =>
                                                                                      current +
                                                                                      ("'" + worksheetName + "'!" +
                                                                                       printArea.RangeAddress.
                                                                                           FirstAddress.ToStringFixed() +
                                                                                       ":" +
                                                                                       printArea.RangeAddress.
                                                                                           LastAddress.ToStringFixed() +
                                                                                       ","));
                    definedName.Text = definedNameText.Substring(0, definedNameText.Length - 1);
                    definedNames.AppendChild(definedName);
                }

                foreach (IXLNamedRange nr in worksheet.NamedRanges)
                {
                    var definedName = new DefinedName
                                          {
                                              Name = nr.Name,
                                              LocalSheetId = sheetId,
                                              Text = nr.ToString()
                                          };
                    if (!StringExtensions.IsNullOrWhiteSpace(nr.Comment))
                        definedName.Comment = nr.Comment;
                    definedNames.AppendChild(definedName);
                }


                string definedNameTextRow = String.Empty;
                string definedNameTextColumn = String.Empty;
                if (worksheet.PageSetup.FirstRowToRepeatAtTop > 0)
                {
                    definedNameTextRow = "'" + worksheet.Name + "'!" + worksheet.PageSetup.FirstRowToRepeatAtTop
                                         + ":" + worksheet.PageSetup.LastRowToRepeatAtTop;
                }
                if (worksheet.PageSetup.FirstColumnToRepeatAtLeft > 0)
                {
                    int minColumn = worksheet.PageSetup.FirstColumnToRepeatAtLeft;
                    int maxColumn = worksheet.PageSetup.LastColumnToRepeatAtLeft;
                    definedNameTextColumn = "'" + worksheet.Name + "'!" +
                                            ExcelHelper.GetColumnLetterFromNumber(minColumn)
                                            + ":" + ExcelHelper.GetColumnLetterFromNumber(maxColumn);
                }

                string titles;
                if (definedNameTextColumn.Length > 0)
                {
                    titles = definedNameTextColumn;
                    if (definedNameTextRow.Length > 0)
                        titles += "," + definedNameTextRow;
                }
                else
                    titles = definedNameTextRow;

                if (titles.Length <= 0) continue;

                var definedName2 = new DefinedName
                                      {
                                          Name = "_xlnm.Print_Titles",
                                          LocalSheetId = sheetId,
                                          Text = titles
                                      };

                definedNames.AppendChild(definedName2);
            }

            foreach (IXLNamedRange nr in NamedRanges)
            {
                var definedName = new DefinedName
                                      {
                                          Name = nr.Name,
                                          Text = nr.ToString()
                                      };
                if (!StringExtensions.IsNullOrWhiteSpace(nr.Comment))
                    definedName.Comment = nr.Comment;
                definedNames.AppendChild(definedName);
            }

            if (workbook.DefinedNames == null)
                workbook.DefinedNames = new DefinedNames();

            foreach (DefinedName dn in definedNames)
            {
                String dnName = dn.Name.Value;
                var dnLocalSheetId = dn.LocalSheetId;
                var existingDefinedName = workbook.DefinedNames
                    .Elements<DefinedName>()
                    .FirstOrDefault(d =>
                                    String.Compare(d.Name.Value, dnName, true) == 0
                                    && (
                                           (d.LocalSheetId != null && dnLocalSheetId != null &&
                                            d.LocalSheetId.InnerText == dnLocalSheetId.InnerText)
                                           || d.LocalSheetId == null
                                           || dnLocalSheetId == null)
                    );
                if (existingDefinedName != null)
                {
                    existingDefinedName.Text = dn.Text;
                    existingDefinedName.LocalSheetId = dn.LocalSheetId;
                    existingDefinedName.Comment = dn.Comment;
                }
                else
                    workbook.DefinedNames.AppendChild(dn.CloneNode(true));
            }

            if (workbook.CalculationProperties == null)
                workbook.CalculationProperties = new CalculationProperties {CalculationId = 125725U};

            if (CalculateMode == XLCalculateMode.Default)
                workbook.CalculationProperties.CalculationMode = null;
            else
                workbook.CalculationProperties.CalculationMode = CalculateMode.ToOpenXml();

            if (ReferenceStyle == XLReferenceStyle.Default)
                workbook.CalculationProperties.ReferenceMode = null;
            else
                workbook.CalculationProperties.ReferenceMode = ReferenceStyle.ToOpenXml();
        }

        private void GenerateSharedStringTablePartContent(SharedStringTablePart sharedStringTablePart, SaveContext context)
        {
            sharedStringTablePart.SharedStringTable = new SharedStringTable {Count = 0, UniqueCount = 0};

            Int32 stringId = 0;

            var newStrings = new Dictionary<String, Int32>();
            var newRichStrings = new Dictionary<IXLRichText, Int32>();
            foreach (XLCell c in Worksheets.Cast<XLWorksheet>().SelectMany(w => w.Internals.CellsCollection.GetCells().Where(c => c.DataType == XLCellValues.Text
                                                                                                                                  && c.ShareString
                                                                                                                                  && c.InnerText.Length > 0)))
            {
                if (c.HasRichText)
                {
                    if (newRichStrings.ContainsKey(c.RichText))
                        c.SharedStringId = newRichStrings[c.RichText];
                    else
                    {
                        var sharedStringItem = new SharedStringItem();
                        foreach (IXLRichString rt in c.RichText)
                        {
                            sharedStringItem.Append(GetRun(rt));
                        }

                        if (c.RichText.HasPhonetics)
                        {
                            foreach (IXLPhonetic p in c.RichText.Phonetics)
                            {
                                var phoneticRun = new PhoneticRun
                                                      {
                                                          BaseTextStartIndex = (UInt32)p.Start,
                                                          EndingBaseIndex = (UInt32)p.End
                                                      };

                                var text = new Text {Text = p.Text};
                                if (p.Text.PreserveSpaces())
                                    text.Space = SpaceProcessingModeValues.Preserve;

                                phoneticRun.Append(text);
                                sharedStringItem.Append(phoneticRun);
                            }
                            var f = new XLFont(null, c.RichText.Phonetics);
                            if (!context.SharedFonts.ContainsKey(f))
                                context.SharedFonts.Add(f, new FontInfo {Font = f});

                            var phoneticProperties = new PhoneticProperties
                                                         {
                                                             FontId =
                                                                 context.SharedFonts[
                                                                     new XLFont(null, c.RichText.Phonetics)].
                                                                 FontId
                                                         };
                            if (c.RichText.Phonetics.Alignment != XLPhoneticAlignment.Left)
                                phoneticProperties.Alignment = c.RichText.Phonetics.Alignment.ToOpenXml();
                            if (c.RichText.Phonetics.Type != XLPhoneticType.FullWidthKatakana)
                                phoneticProperties.Type = c.RichText.Phonetics.Type.ToOpenXml();

                            sharedStringItem.Append(phoneticProperties);
                        }

                        sharedStringTablePart.SharedStringTable.Append(sharedStringItem);
                        sharedStringTablePart.SharedStringTable.Count += 1;
                        sharedStringTablePart.SharedStringTable.UniqueCount += 1;

                        newRichStrings.Add(c.RichText, stringId);
                        c.SharedStringId = stringId;

                        stringId++;
                    }
                }
                else
                {
                    if (newStrings.ContainsKey(c.Value.ToString()))
                        c.SharedStringId = newStrings[c.Value.ToString()];
                    else
                    {
                        String s = c.Value.ToString();
                        var sharedStringItem = new SharedStringItem();
                        var text = new Text {Text = s};
                        if (s.StartsWith(" ") || s.EndsWith(" "))
                            text.Space = SpaceProcessingModeValues.Preserve;
                        sharedStringItem.Append(text);
                        sharedStringTablePart.SharedStringTable.Append(sharedStringItem);
                        sharedStringTablePart.SharedStringTable.Count += 1;
                        sharedStringTablePart.SharedStringTable.UniqueCount += 1;

                        newStrings.Add(c.Value.ToString(), stringId);
                        c.SharedStringId = stringId;

                        stringId++;
                    }
                }
            }
        }

        private static DocumentFormat.OpenXml.Spreadsheet.Run GetRun(IXLRichString rt)
        {
            var run = new DocumentFormat.OpenXml.Spreadsheet.Run();

            var runProperties = new DocumentFormat.OpenXml.Spreadsheet.RunProperties();

            var bold = rt.Bold ? new Bold() : null;
            var italic = rt.Italic ? new Italic() : null;
            var underline = rt.Underline != XLFontUnderlineValues.None
                                ? new Underline {Val = rt.Underline.ToOpenXml()}
                                : null;
            var strike = rt.Strikethrough ? new Strike() : null;
            var verticalAlignment = new VerticalTextAlignment
                                        {Val = rt.VerticalAlignment.ToOpenXml()};
            var shadow = rt.Shadow ? new Shadow() : null;
            var fontSize = new FontSize {Val = rt.FontSize};
            var color = GetNewColor(rt.FontColor);
            var fontName = new RunFont {Val = rt.FontName};
            var fontFamilyNumbering = new FontFamily {Val = (Int32)rt.FontFamilyNumbering};

            if (bold != null) runProperties.Append(bold);
            if (italic != null) runProperties.Append(italic);

            if (strike != null) runProperties.Append(strike);
            if (shadow != null) runProperties.Append(shadow);
            if (underline != null) runProperties.Append(underline);
            runProperties.Append(verticalAlignment);

            runProperties.Append(fontSize);
            runProperties.Append(color);
            runProperties.Append(fontName);
            runProperties.Append(fontFamilyNumbering);

            var text = new Text {Text = rt.Text};
            if (rt.Text.PreserveSpaces())
                text.Space = SpaceProcessingModeValues.Preserve;

            run.Append(runProperties);
            run.Append(text);
            return run;
        }

        private void GenerateCalculationChainPartContent(WorkbookPart workbookPart, SaveContext context)
        {
            string thisRelId = context.RelIdGenerator.GetNext(RelType.Workbook);
            if (workbookPart.CalculationChainPart == null)
                workbookPart.AddNewPart<CalculationChainPart>(thisRelId);

            if (workbookPart.CalculationChainPart.CalculationChain == null)
                workbookPart.CalculationChainPart.CalculationChain = new CalculationChain();

            var calculationChain = workbookPart.CalculationChainPart.CalculationChain;
            calculationChain.RemoveAllChildren<CalculationCell>();
            //var calculationCells = new Dictionary<String, List<CalculationCell>>();
            //foreach(var calculationCell in calculationChain.Elements<CalculationCell>().Where(cc => cc.CellReference != null))
            //{
            //    String cellReference = calculationCell.CellReference.Value;
            //    if (!calculationCells.ContainsKey(cellReference))
            //        calculationCells.Add(cellReference, new List<CalculationCell>());
            //    calculationCell.
            //    calculationCells[cellReference].Add(calculationCell);
            //}

            foreach (XLWorksheet worksheet in WorksheetsInternal)
            {
                var cellsWithoutFormulas = new HashSet<String>();
                foreach (XLCell c in worksheet.Internals.CellsCollection.GetCells())
                {
                    if (StringExtensions.IsNullOrWhiteSpace(c.FormulaA1))
                        cellsWithoutFormulas.Add(c.Address.ToStringRelative());
                    else
                    {
                        //var calculationCells = calculationChain.Elements<CalculationCell>().Where(
                        //    cc => cc.CellReference != null && cc.CellReference == c.Address.ToString()).Select(cc => cc).ToList();

                        //calculationCells.ForEach(cc => calculationChain.RemoveChild(cc));

                        if (c.FormulaA1.StartsWith("{"))
                        {
                            calculationChain.AppendChild(new CalculationCell
                                                             {
                                                                 CellReference = c.Address.ToString(),
                                                                 SheetId = worksheet.SheetId,
                                                                 Array = true
                                                             });
                            calculationChain.AppendChild(new CalculationCell
                                                             {CellReference = c.Address.ToString(), InChildChain = true});
                        }
                        else
                        {
                            calculationChain.AppendChild(new CalculationCell
                                                             {
                                                                 CellReference = c.Address.ToString(),
                                                                 SheetId = worksheet.SheetId
                                                             });
                        }
                    }
                }

                //var cCellsToRemove = new List<CalculationCell>();
                var m = from cc in calculationChain.Elements<CalculationCell>()
                        where !(cc.SheetId != null || cc.InChildChain != null)
                              && calculationChain.Elements<CalculationCell>()
                                     .Where(c1 => c1.SheetId != null)
                                     .Select(c1 => c1.CellReference.Value)
                                     .Contains(cc.CellReference.Value)
                              || cellsWithoutFormulas.Contains(cc.CellReference.Value)
                        select cc;
                //m.ToList().ForEach(cc => cCellsToRemove.Add(cc));
                m.ToList().ForEach(cc => calculationChain.RemoveChild(cc));
            }

            if (!calculationChain.Any())
                workbookPart.DeletePart(workbookPart.CalculationChainPart);
        }

        private void GenerateThemePartContent(ThemePart themePart)
        {
            var theme1 = new Theme {Name = "Office Theme"};
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            var themeElements1 = new ThemeElements();

            var colorScheme1 = new ColorScheme {Name = "Office"};

            var dark1Color1 = new Dark1Color();
            var systemColor1 = new SystemColor
                                   {
                                       Val = SystemColorValues.WindowText,
                                       LastColor = Theme.Text1.Color.ToHex().Substring(2)
                                   };

            dark1Color1.AppendChild(systemColor1);

            var light1Color1 = new Light1Color();
            var systemColor2 = new SystemColor
                                   {
                                       Val = SystemColorValues.Window,
                                       LastColor = Theme.Background1.Color.ToHex().Substring(2)
                                   };

            light1Color1.AppendChild(systemColor2);

            var dark2Color1 = new Dark2Color();
            var rgbColorModelHex1 = new RgbColorModelHex {Val = Theme.Text2.Color.ToHex().Substring(2)};

            dark2Color1.AppendChild(rgbColorModelHex1);

            var light2Color1 = new Light2Color();
            var rgbColorModelHex2 = new RgbColorModelHex {Val = Theme.Background2.Color.ToHex().Substring(2)};

            light2Color1.AppendChild(rgbColorModelHex2);

            var accent1Color1 = new Accent1Color();
            var rgbColorModelHex3 = new RgbColorModelHex {Val = Theme.Accent1.Color.ToHex().Substring(2)};

            accent1Color1.AppendChild(rgbColorModelHex3);

            var accent2Color1 = new Accent2Color();
            var rgbColorModelHex4 = new RgbColorModelHex {Val = Theme.Accent2.Color.ToHex().Substring(2)};

            accent2Color1.AppendChild(rgbColorModelHex4);

            var accent3Color1 = new Accent3Color();
            var rgbColorModelHex5 = new RgbColorModelHex {Val = Theme.Accent3.Color.ToHex().Substring(2)};

            accent3Color1.AppendChild(rgbColorModelHex5);

            var accent4Color1 = new Accent4Color();
            var rgbColorModelHex6 = new RgbColorModelHex {Val = Theme.Accent4.Color.ToHex().Substring(2)};

            accent4Color1.AppendChild(rgbColorModelHex6);

            var accent5Color1 = new Accent5Color();
            var rgbColorModelHex7 = new RgbColorModelHex {Val = Theme.Accent5.Color.ToHex().Substring(2)};

            accent5Color1.AppendChild(rgbColorModelHex7);

            var accent6Color1 = new Accent6Color();
            var rgbColorModelHex8 = new RgbColorModelHex {Val = Theme.Accent6.Color.ToHex().Substring(2)};

            accent6Color1.AppendChild(rgbColorModelHex8);

            var hyperlink1 = new DocumentFormat.OpenXml.Drawing.Hyperlink();
            var rgbColorModelHex9 = new RgbColorModelHex {Val = Theme.Hyperlink.Color.ToHex().Substring(2)};

            hyperlink1.AppendChild(rgbColorModelHex9);

            var followedHyperlinkColor1 = new FollowedHyperlinkColor();
            var rgbColorModelHex10 = new RgbColorModelHex {Val = Theme.FollowedHyperlink.Color.ToHex().Substring(2)};

            followedHyperlinkColor1.AppendChild(rgbColorModelHex10);

            colorScheme1.AppendChild(dark1Color1);
            colorScheme1.AppendChild(light1Color1);
            colorScheme1.AppendChild(dark2Color1);
            colorScheme1.AppendChild(light2Color1);
            colorScheme1.AppendChild(accent1Color1);
            colorScheme1.AppendChild(accent2Color1);
            colorScheme1.AppendChild(accent3Color1);
            colorScheme1.AppendChild(accent4Color1);
            colorScheme1.AppendChild(accent5Color1);
            colorScheme1.AppendChild(accent6Color1);
            colorScheme1.AppendChild(hyperlink1);
            colorScheme1.AppendChild(followedHyperlinkColor1);

            var fontScheme2 = new FontScheme {Name = "Office"};

            var majorFont1 = new MajorFont();
            var latinFont1 = new LatinFont {Typeface = "Cambria"};
            var eastAsianFont1 = new EastAsianFont {Typeface = ""};
            var complexScriptFont1 = new ComplexScriptFont {Typeface = ""};
            var supplementalFont1 = new SupplementalFont {Script = "Jpan", Typeface = "ＭＳ Ｐゴシック"};
            var supplementalFont2 = new SupplementalFont {Script = "Hang", Typeface = "맑은 고딕"};
            var supplementalFont3 = new SupplementalFont {Script = "Hans", Typeface = "宋体"};
            var supplementalFont4 = new SupplementalFont {Script = "Hant", Typeface = "新細明體"};
            var supplementalFont5 = new SupplementalFont {Script = "Arab", Typeface = "Times New Roman"};
            var supplementalFont6 = new SupplementalFont {Script = "Hebr", Typeface = "Times New Roman"};
            var supplementalFont7 = new SupplementalFont {Script = "Thai", Typeface = "Tahoma"};
            var supplementalFont8 = new SupplementalFont {Script = "Ethi", Typeface = "Nyala"};
            var supplementalFont9 = new SupplementalFont {Script = "Beng", Typeface = "Vrinda"};
            var supplementalFont10 = new SupplementalFont {Script = "Gujr", Typeface = "Shruti"};
            var supplementalFont11 = new SupplementalFont {Script = "Khmr", Typeface = "MoolBoran"};
            var supplementalFont12 = new SupplementalFont {Script = "Knda", Typeface = "Tunga"};
            var supplementalFont13 = new SupplementalFont {Script = "Guru", Typeface = "Raavi"};
            var supplementalFont14 = new SupplementalFont {Script = "Cans", Typeface = "Euphemia"};
            var supplementalFont15 = new SupplementalFont {Script = "Cher", Typeface = "Plantagenet Cherokee"};
            var supplementalFont16 = new SupplementalFont {Script = "Yiii", Typeface = "Microsoft Yi Baiti"};
            var supplementalFont17 = new SupplementalFont {Script = "Tibt", Typeface = "Microsoft Himalaya"};
            var supplementalFont18 = new SupplementalFont {Script = "Thaa", Typeface = "MV Boli"};
            var supplementalFont19 = new SupplementalFont {Script = "Deva", Typeface = "Mangal"};
            var supplementalFont20 = new SupplementalFont {Script = "Telu", Typeface = "Gautami"};
            var supplementalFont21 = new SupplementalFont {Script = "Taml", Typeface = "Latha"};
            var supplementalFont22 = new SupplementalFont {Script = "Syrc", Typeface = "Estrangelo Edessa"};
            var supplementalFont23 = new SupplementalFont {Script = "Orya", Typeface = "Kalinga"};
            var supplementalFont24 = new SupplementalFont {Script = "Mlym", Typeface = "Kartika"};
            var supplementalFont25 = new SupplementalFont {Script = "Laoo", Typeface = "DokChampa"};
            var supplementalFont26 = new SupplementalFont {Script = "Sinh", Typeface = "Iskoola Pota"};
            var supplementalFont27 = new SupplementalFont {Script = "Mong", Typeface = "Mongolian Baiti"};
            var supplementalFont28 = new SupplementalFont {Script = "Viet", Typeface = "Times New Roman"};
            var supplementalFont29 = new SupplementalFont {Script = "Uigh", Typeface = "Microsoft Uighur"};

            majorFont1.AppendChild(latinFont1);
            majorFont1.AppendChild(eastAsianFont1);
            majorFont1.AppendChild(complexScriptFont1);
            majorFont1.AppendChild(supplementalFont1);
            majorFont1.AppendChild(supplementalFont2);
            majorFont1.AppendChild(supplementalFont3);
            majorFont1.AppendChild(supplementalFont4);
            majorFont1.AppendChild(supplementalFont5);
            majorFont1.AppendChild(supplementalFont6);
            majorFont1.AppendChild(supplementalFont7);
            majorFont1.AppendChild(supplementalFont8);
            majorFont1.AppendChild(supplementalFont9);
            majorFont1.AppendChild(supplementalFont10);
            majorFont1.AppendChild(supplementalFont11);
            majorFont1.AppendChild(supplementalFont12);
            majorFont1.AppendChild(supplementalFont13);
            majorFont1.AppendChild(supplementalFont14);
            majorFont1.AppendChild(supplementalFont15);
            majorFont1.AppendChild(supplementalFont16);
            majorFont1.AppendChild(supplementalFont17);
            majorFont1.AppendChild(supplementalFont18);
            majorFont1.AppendChild(supplementalFont19);
            majorFont1.AppendChild(supplementalFont20);
            majorFont1.AppendChild(supplementalFont21);
            majorFont1.AppendChild(supplementalFont22);
            majorFont1.AppendChild(supplementalFont23);
            majorFont1.AppendChild(supplementalFont24);
            majorFont1.AppendChild(supplementalFont25);
            majorFont1.AppendChild(supplementalFont26);
            majorFont1.AppendChild(supplementalFont27);
            majorFont1.AppendChild(supplementalFont28);
            majorFont1.AppendChild(supplementalFont29);

            var minorFont1 = new MinorFont();
            var latinFont2 = new LatinFont {Typeface = "Calibri"};
            var eastAsianFont2 = new EastAsianFont {Typeface = ""};
            var complexScriptFont2 = new ComplexScriptFont {Typeface = ""};
            var supplementalFont30 = new SupplementalFont {Script = "Jpan", Typeface = "ＭＳ Ｐゴシック"};
            var supplementalFont31 = new SupplementalFont {Script = "Hang", Typeface = "맑은 고딕"};
            var supplementalFont32 = new SupplementalFont {Script = "Hans", Typeface = "宋体"};
            var supplementalFont33 = new SupplementalFont {Script = "Hant", Typeface = "新細明體"};
            var supplementalFont34 = new SupplementalFont {Script = "Arab", Typeface = "Arial"};
            var supplementalFont35 = new SupplementalFont {Script = "Hebr", Typeface = "Arial"};
            var supplementalFont36 = new SupplementalFont {Script = "Thai", Typeface = "Tahoma"};
            var supplementalFont37 = new SupplementalFont {Script = "Ethi", Typeface = "Nyala"};
            var supplementalFont38 = new SupplementalFont {Script = "Beng", Typeface = "Vrinda"};
            var supplementalFont39 = new SupplementalFont {Script = "Gujr", Typeface = "Shruti"};
            var supplementalFont40 = new SupplementalFont {Script = "Khmr", Typeface = "DaunPenh"};
            var supplementalFont41 = new SupplementalFont {Script = "Knda", Typeface = "Tunga"};
            var supplementalFont42 = new SupplementalFont {Script = "Guru", Typeface = "Raavi"};
            var supplementalFont43 = new SupplementalFont {Script = "Cans", Typeface = "Euphemia"};
            var supplementalFont44 = new SupplementalFont {Script = "Cher", Typeface = "Plantagenet Cherokee"};
            var supplementalFont45 = new SupplementalFont {Script = "Yiii", Typeface = "Microsoft Yi Baiti"};
            var supplementalFont46 = new SupplementalFont {Script = "Tibt", Typeface = "Microsoft Himalaya"};
            var supplementalFont47 = new SupplementalFont {Script = "Thaa", Typeface = "MV Boli"};
            var supplementalFont48 = new SupplementalFont {Script = "Deva", Typeface = "Mangal"};
            var supplementalFont49 = new SupplementalFont {Script = "Telu", Typeface = "Gautami"};
            var supplementalFont50 = new SupplementalFont {Script = "Taml", Typeface = "Latha"};
            var supplementalFont51 = new SupplementalFont {Script = "Syrc", Typeface = "Estrangelo Edessa"};
            var supplementalFont52 = new SupplementalFont {Script = "Orya", Typeface = "Kalinga"};
            var supplementalFont53 = new SupplementalFont {Script = "Mlym", Typeface = "Kartika"};
            var supplementalFont54 = new SupplementalFont {Script = "Laoo", Typeface = "DokChampa"};
            var supplementalFont55 = new SupplementalFont {Script = "Sinh", Typeface = "Iskoola Pota"};
            var supplementalFont56 = new SupplementalFont {Script = "Mong", Typeface = "Mongolian Baiti"};
            var supplementalFont57 = new SupplementalFont {Script = "Viet", Typeface = "Arial"};
            var supplementalFont58 = new SupplementalFont {Script = "Uigh", Typeface = "Microsoft Uighur"};

            minorFont1.AppendChild(latinFont2);
            minorFont1.AppendChild(eastAsianFont2);
            minorFont1.AppendChild(complexScriptFont2);
            minorFont1.AppendChild(supplementalFont30);
            minorFont1.AppendChild(supplementalFont31);
            minorFont1.AppendChild(supplementalFont32);
            minorFont1.AppendChild(supplementalFont33);
            minorFont1.AppendChild(supplementalFont34);
            minorFont1.AppendChild(supplementalFont35);
            minorFont1.AppendChild(supplementalFont36);
            minorFont1.AppendChild(supplementalFont37);
            minorFont1.AppendChild(supplementalFont38);
            minorFont1.AppendChild(supplementalFont39);
            minorFont1.AppendChild(supplementalFont40);
            minorFont1.AppendChild(supplementalFont41);
            minorFont1.AppendChild(supplementalFont42);
            minorFont1.AppendChild(supplementalFont43);
            minorFont1.AppendChild(supplementalFont44);
            minorFont1.AppendChild(supplementalFont45);
            minorFont1.AppendChild(supplementalFont46);
            minorFont1.AppendChild(supplementalFont47);
            minorFont1.AppendChild(supplementalFont48);
            minorFont1.AppendChild(supplementalFont49);
            minorFont1.AppendChild(supplementalFont50);
            minorFont1.AppendChild(supplementalFont51);
            minorFont1.AppendChild(supplementalFont52);
            minorFont1.AppendChild(supplementalFont53);
            minorFont1.AppendChild(supplementalFont54);
            minorFont1.AppendChild(supplementalFont55);
            minorFont1.AppendChild(supplementalFont56);
            minorFont1.AppendChild(supplementalFont57);
            minorFont1.AppendChild(supplementalFont58);

            fontScheme2.AppendChild(majorFont1);
            fontScheme2.AppendChild(minorFont1);

            var formatScheme1 = new FormatScheme {Name = "Office"};

            var fillStyleList1 = new FillStyleList();

            var solidFill1 = new SolidFill();
            var schemeColor1 = new SchemeColor {Val = SchemeColorValues.PhColor};

            solidFill1.AppendChild(schemeColor1);

            var gradientFill1 = new GradientFill {RotateWithShape = true};

            var gradientStopList1 = new GradientStopList();

            var gradientStop1 = new GradientStop {Position = 0};

            var schemeColor2 = new SchemeColor {Val = SchemeColorValues.PhColor};
            var tint1 = new Tint {Val = 50000};
            var saturationModulation1 = new SaturationModulation {Val = 300000};

            schemeColor2.AppendChild(tint1);
            schemeColor2.AppendChild(saturationModulation1);

            gradientStop1.AppendChild(schemeColor2);

            var gradientStop2 = new GradientStop {Position = 35000};

            var schemeColor3 = new SchemeColor {Val = SchemeColorValues.PhColor};
            var tint2 = new Tint {Val = 37000};
            var saturationModulation2 = new SaturationModulation {Val = 300000};

            schemeColor3.AppendChild(tint2);
            schemeColor3.AppendChild(saturationModulation2);

            gradientStop2.AppendChild(schemeColor3);

            var gradientStop3 = new GradientStop {Position = 100000};

            var schemeColor4 = new SchemeColor {Val = SchemeColorValues.PhColor};
            var tint3 = new Tint {Val = 15000};
            var saturationModulation3 = new SaturationModulation {Val = 350000};

            schemeColor4.AppendChild(tint3);
            schemeColor4.AppendChild(saturationModulation3);

            gradientStop3.AppendChild(schemeColor4);

            gradientStopList1.AppendChild(gradientStop1);
            gradientStopList1.AppendChild(gradientStop2);
            gradientStopList1.AppendChild(gradientStop3);
            var linearGradientFill1 = new LinearGradientFill {Angle = 16200000, Scaled = true};

            gradientFill1.AppendChild(gradientStopList1);
            gradientFill1.AppendChild(linearGradientFill1);

            var gradientFill2 = new GradientFill {RotateWithShape = true};

            var gradientStopList2 = new GradientStopList();

            var gradientStop4 = new GradientStop {Position = 0};

            var schemeColor5 = new SchemeColor {Val = SchemeColorValues.PhColor};
            var shade1 = new Shade {Val = 51000};
            var saturationModulation4 = new SaturationModulation {Val = 130000};

            schemeColor5.AppendChild(shade1);
            schemeColor5.AppendChild(saturationModulation4);

            gradientStop4.AppendChild(schemeColor5);

            var gradientStop5 = new GradientStop {Position = 80000};

            var schemeColor6 = new SchemeColor {Val = SchemeColorValues.PhColor};
            var shade2 = new Shade {Val = 93000};
            var saturationModulation5 = new SaturationModulation {Val = 130000};

            schemeColor6.AppendChild(shade2);
            schemeColor6.AppendChild(saturationModulation5);

            gradientStop5.AppendChild(schemeColor6);

            var gradientStop6 = new GradientStop {Position = 100000};

            var schemeColor7 = new SchemeColor {Val = SchemeColorValues.PhColor};
            var shade3 = new Shade {Val = 94000};
            var saturationModulation6 = new SaturationModulation {Val = 135000};

            schemeColor7.AppendChild(shade3);
            schemeColor7.AppendChild(saturationModulation6);

            gradientStop6.AppendChild(schemeColor7);

            gradientStopList2.AppendChild(gradientStop4);
            gradientStopList2.AppendChild(gradientStop5);
            gradientStopList2.AppendChild(gradientStop6);
            var linearGradientFill2 = new LinearGradientFill {Angle = 16200000, Scaled = false};

            gradientFill2.AppendChild(gradientStopList2);
            gradientFill2.AppendChild(linearGradientFill2);

            fillStyleList1.AppendChild(solidFill1);
            fillStyleList1.AppendChild(gradientFill1);
            fillStyleList1.AppendChild(gradientFill2);

            var lineStyleList1 = new LineStyleList();

            var outline1 = new Outline
                               {
                                   Width = 9525,
                                   CapType = LineCapValues.Flat,
                                   CompoundLineType = CompoundLineValues.Single,
                                   Alignment = PenAlignmentValues.Center
                               };

            var solidFill2 = new SolidFill();

            var schemeColor8 = new SchemeColor {Val = SchemeColorValues.PhColor};
            var shade4 = new Shade {Val = 95000};
            var saturationModulation7 = new SaturationModulation {Val = 105000};

            schemeColor8.AppendChild(shade4);
            schemeColor8.AppendChild(saturationModulation7);

            solidFill2.AppendChild(schemeColor8);
            var presetDash1 = new PresetDash {Val = PresetLineDashValues.Solid};

            outline1.AppendChild(solidFill2);
            outline1.AppendChild(presetDash1);

            var outline2 = new Outline
                               {
                                   Width = 25400,
                                   CapType = LineCapValues.Flat,
                                   CompoundLineType = CompoundLineValues.Single,
                                   Alignment = PenAlignmentValues.Center
                               };

            var solidFill3 = new SolidFill();
            var schemeColor9 = new SchemeColor {Val = SchemeColorValues.PhColor};

            solidFill3.AppendChild(schemeColor9);
            var presetDash2 = new PresetDash {Val = PresetLineDashValues.Solid};

            outline2.AppendChild(solidFill3);
            outline2.AppendChild(presetDash2);

            var outline3 = new Outline
                               {
                                   Width = 38100,
                                   CapType = LineCapValues.Flat,
                                   CompoundLineType = CompoundLineValues.Single,
                                   Alignment = PenAlignmentValues.Center
                               };

            var solidFill4 = new SolidFill();
            var schemeColor10 = new SchemeColor {Val = SchemeColorValues.PhColor};

            solidFill4.AppendChild(schemeColor10);
            var presetDash3 = new PresetDash {Val = PresetLineDashValues.Solid};

            outline3.AppendChild(solidFill4);
            outline3.AppendChild(presetDash3);

            lineStyleList1.AppendChild(outline1);
            lineStyleList1.AppendChild(outline2);
            lineStyleList1.AppendChild(outline3);

            var effectStyleList1 = new EffectStyleList();

            var effectStyle1 = new EffectStyle();

            var effectList1 = new EffectList();

            var outerShadow1 = new OuterShadow
                                   {
                                       BlurRadius = 40000L,
                                       Distance = 20000L,
                                       Direction = 5400000,
                                       RotateWithShape = false
                                   };

            var rgbColorModelHex11 = new RgbColorModelHex {Val = "000000"};
            var alpha1 = new Alpha {Val = 38000};

            rgbColorModelHex11.AppendChild(alpha1);

            outerShadow1.AppendChild(rgbColorModelHex11);

            effectList1.AppendChild(outerShadow1);

            effectStyle1.AppendChild(effectList1);

            var effectStyle2 = new EffectStyle();

            var effectList2 = new EffectList();

            var outerShadow2 = new OuterShadow
                                   {
                                       BlurRadius = 40000L,
                                       Distance = 23000L,
                                       Direction = 5400000,
                                       RotateWithShape = false
                                   };

            var rgbColorModelHex12 = new RgbColorModelHex {Val = "000000"};
            var alpha2 = new Alpha {Val = 35000};

            rgbColorModelHex12.AppendChild(alpha2);

            outerShadow2.AppendChild(rgbColorModelHex12);

            effectList2.AppendChild(outerShadow2);

            effectStyle2.AppendChild(effectList2);

            var effectStyle3 = new EffectStyle();

            var effectList3 = new EffectList();

            var outerShadow3 = new OuterShadow
                                   {
                                       BlurRadius = 40000L,
                                       Distance = 23000L,
                                       Direction = 5400000,
                                       RotateWithShape = false
                                   };

            var rgbColorModelHex13 = new RgbColorModelHex {Val = "000000"};
            var alpha3 = new Alpha {Val = 35000};

            rgbColorModelHex13.AppendChild(alpha3);

            outerShadow3.AppendChild(rgbColorModelHex13);

            effectList3.AppendChild(outerShadow3);

            var scene3DType1 = new Scene3DType();

            var camera1 = new Camera {Preset = PresetCameraValues.OrthographicFront};
            var rotation1 = new Rotation {Latitude = 0, Longitude = 0, Revolution = 0};

            camera1.AppendChild(rotation1);

            var lightRig1 = new LightRig {Rig = LightRigValues.ThreePoints, Direction = LightRigDirectionValues.Top};
            var rotation2 = new Rotation {Latitude = 0, Longitude = 0, Revolution = 1200000};

            lightRig1.AppendChild(rotation2);

            scene3DType1.AppendChild(camera1);
            scene3DType1.AppendChild(lightRig1);

            var shape3DType1 = new Shape3DType();
            var bevelTop1 = new BevelTop {Width = 63500L, Height = 25400L};

            shape3DType1.AppendChild(bevelTop1);

            effectStyle3.AppendChild(effectList3);
            effectStyle3.AppendChild(scene3DType1);
            effectStyle3.AppendChild(shape3DType1);

            effectStyleList1.AppendChild(effectStyle1);
            effectStyleList1.AppendChild(effectStyle2);
            effectStyleList1.AppendChild(effectStyle3);

            var backgroundFillStyleList1 = new BackgroundFillStyleList();

            var solidFill5 = new SolidFill();
            var schemeColor11 = new SchemeColor {Val = SchemeColorValues.PhColor};

            solidFill5.AppendChild(schemeColor11);

            var gradientFill3 = new GradientFill {RotateWithShape = true};

            var gradientStopList3 = new GradientStopList();

            var gradientStop7 = new GradientStop {Position = 0};

            var schemeColor12 = new SchemeColor {Val = SchemeColorValues.PhColor};
            var tint4 = new Tint {Val = 40000};
            var saturationModulation8 = new SaturationModulation {Val = 350000};

            schemeColor12.AppendChild(tint4);
            schemeColor12.AppendChild(saturationModulation8);

            gradientStop7.AppendChild(schemeColor12);

            var gradientStop8 = new GradientStop {Position = 40000};

            var schemeColor13 = new SchemeColor {Val = SchemeColorValues.PhColor};
            var tint5 = new Tint {Val = 45000};
            var shade5 = new Shade {Val = 99000};
            var saturationModulation9 = new SaturationModulation {Val = 350000};

            schemeColor13.AppendChild(tint5);
            schemeColor13.AppendChild(shade5);
            schemeColor13.AppendChild(saturationModulation9);

            gradientStop8.AppendChild(schemeColor13);

            var gradientStop9 = new GradientStop {Position = 100000};

            var schemeColor14 = new SchemeColor {Val = SchemeColorValues.PhColor};
            var shade6 = new Shade {Val = 20000};
            var saturationModulation10 = new SaturationModulation {Val = 255000};

            schemeColor14.AppendChild(shade6);
            schemeColor14.AppendChild(saturationModulation10);

            gradientStop9.AppendChild(schemeColor14);

            gradientStopList3.AppendChild(gradientStop7);
            gradientStopList3.AppendChild(gradientStop8);
            gradientStopList3.AppendChild(gradientStop9);

            var pathGradientFill1 = new PathGradientFill {Path = PathShadeValues.Circle};
            var fillToRectangle1 = new FillToRectangle {Left = 50000, Top = -80000, Right = 50000, Bottom = 180000};

            pathGradientFill1.AppendChild(fillToRectangle1);

            gradientFill3.AppendChild(gradientStopList3);
            gradientFill3.AppendChild(pathGradientFill1);

            var gradientFill4 = new GradientFill {RotateWithShape = true};

            var gradientStopList4 = new GradientStopList();

            var gradientStop10 = new GradientStop {Position = 0};

            var schemeColor15 = new SchemeColor {Val = SchemeColorValues.PhColor};
            var tint6 = new Tint {Val = 80000};
            var saturationModulation11 = new SaturationModulation {Val = 300000};

            schemeColor15.AppendChild(tint6);
            schemeColor15.AppendChild(saturationModulation11);

            gradientStop10.AppendChild(schemeColor15);

            var gradientStop11 = new GradientStop {Position = 100000};

            var schemeColor16 = new SchemeColor {Val = SchemeColorValues.PhColor};
            var shade7 = new Shade {Val = 30000};
            var saturationModulation12 = new SaturationModulation {Val = 200000};

            schemeColor16.AppendChild(shade7);
            schemeColor16.AppendChild(saturationModulation12);

            gradientStop11.AppendChild(schemeColor16);

            gradientStopList4.AppendChild(gradientStop10);
            gradientStopList4.AppendChild(gradientStop11);

            var pathGradientFill2 = new PathGradientFill {Path = PathShadeValues.Circle};
            var fillToRectangle2 = new FillToRectangle {Left = 50000, Top = 50000, Right = 50000, Bottom = 50000};

            pathGradientFill2.AppendChild(fillToRectangle2);

            gradientFill4.AppendChild(gradientStopList4);
            gradientFill4.AppendChild(pathGradientFill2);

            backgroundFillStyleList1.AppendChild(solidFill5);
            backgroundFillStyleList1.AppendChild(gradientFill3);
            backgroundFillStyleList1.AppendChild(gradientFill4);

            formatScheme1.AppendChild(fillStyleList1);
            formatScheme1.AppendChild(lineStyleList1);
            formatScheme1.AppendChild(effectStyleList1);
            formatScheme1.AppendChild(backgroundFillStyleList1);

            themeElements1.AppendChild(colorScheme1);
            themeElements1.AppendChild(fontScheme2);
            themeElements1.AppendChild(formatScheme1);
            var objectDefaults1 = new ObjectDefaults();
            var extraColorSchemeList1 = new ExtraColorSchemeList();

            theme1.AppendChild(themeElements1);
            theme1.AppendChild(objectDefaults1);
            theme1.AppendChild(extraColorSchemeList1);

            themePart.Theme = theme1;
        }

        private void GenerateCustomFilePropertiesPartContent(CustomFilePropertiesPart customFilePropertiesPart1)
        {
            var properties2 = new DocumentFormat.OpenXml.CustomProperties.Properties();
            properties2.AddNamespaceDeclaration("vt",
                                                "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Int32 propertyId = 1;
            foreach (IXLCustomProperty p in CustomProperties)
            {
                propertyId++;
                var customDocumentProperty = new CustomDocumentProperty
                                                 {
                                                     FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
                                                     PropertyId = propertyId,
                                                     Name = p.Name
                                                 };
                if (p.Type == XLCustomPropertyType.Text)
                {
                    var vTlpwstr1 = new VTLPWSTR {Text = p.GetValue<string>()};
                    customDocumentProperty.AppendChild(vTlpwstr1);
                }
                else if (p.Type == XLCustomPropertyType.Date)
                {
                    var vTFileTime1 = new VTFileTime
                                          {
                                              Text =
                                                  p.GetValue<DateTime>().ToUniversalTime().ToString(
                                                      "yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'")
                                          };
                    customDocumentProperty.AppendChild(vTFileTime1);
                }
                else if (p.Type == XLCustomPropertyType.Number)
                {
                    var vTDouble1 = new VTDouble
                                        {
                                            Text = p.GetValue<Double>().ToString(CultureInfo.InvariantCulture)
                                        };
                    customDocumentProperty.AppendChild(vTDouble1);
                }
                else
                {
                    var vTBool1 = new VTBool {Text = p.GetValue<Boolean>().ToString().ToLower()};
                    customDocumentProperty.AppendChild(vTBool1);
                }
                properties2.AppendChild(customDocumentProperty);
            }

            customFilePropertiesPart1.Properties = properties2;
        }

        private void SetPackageProperties(OpenXmlPackage document)
        {
            var created = Properties.Created == DateTime.MinValue ? DateTime.Now : Properties.Created;
            var modified = Properties.Modified == DateTime.MinValue ? DateTime.Now : Properties.Modified;
            document.PackageProperties.Created = created;
            document.PackageProperties.Modified = modified;
            document.PackageProperties.LastModifiedBy = Properties.LastModifiedBy;

            document.PackageProperties.Creator = Properties.Author;
            document.PackageProperties.Title = Properties.Title;
            document.PackageProperties.Subject = Properties.Subject;
            document.PackageProperties.Category = Properties.Category;
            document.PackageProperties.Keywords = Properties.Keywords;
            document.PackageProperties.Description = Properties.Comments;
            document.PackageProperties.ContentStatus = Properties.Status;
        }

        private static string GetTableName(String originalTableName, SaveContext context)
        {
            string tableName = originalTableName.RemoveSpecialCharacters();
            string name = tableName;
            if (context.TableNames.Contains(name))
            {
                Int32 i = 1;
                name = tableName + i.ToStringLookup();
                while (context.TableNames.Contains(name))
                {
                    i++;
                    name = tableName + i.ToStringLookup();
                }
            }

            context.TableNames.Add(name);
            return name;
        }

        private static void GenerateTableDefinitionPartContent(TableDefinitionPart tableDefinitionPart, XLTable xlTable,
                                                               SaveContext context)
        {
            context.TableId++;
            string reference = xlTable.RangeAddress.FirstAddress + ":" + xlTable.RangeAddress.LastAddress;
            String tableName = GetTableName(xlTable.Name, context);
            var table = new Table
                            {
                                Id = context.TableId,
                                Name = tableName,
                                DisplayName = tableName,
                                Reference = reference
                            };

            if (xlTable.ShowTotalsRow)
                table.TotalsRowCount = 1;
            else
                table.TotalsRowShown = false;

            var tableColumns1 = new TableColumns {Count = (UInt32)xlTable.ColumnCount()};
            UInt32 columnId = 0;
            foreach (IXLCell cell in xlTable.HeadersRow().Cells())
            {
                columnId++;
                String fieldName = cell.GetString();
                var xlField = xlTable.Field(fieldName);
                var tableColumn1 = new TableColumn
                                       {
                                           Id = columnId,
                                           Name = fieldName
                                       };
                if (xlTable.ShowTotalsRow)
                {
                    if (xlField.TotalsRowFunction != XLTotalsRowFunction.None)
                    {
                        tableColumn1.TotalsRowFunction = xlField.TotalsRowFunction.ToOpenXml();

                        if (xlField.TotalsRowFunction == XLTotalsRowFunction.Custom)
                            tableColumn1.TotalsRowFormula = new TotalsRowFormula(xlField.TotalsRowFormulaA1);
                    }

                    if (!StringExtensions.IsNullOrWhiteSpace(xlField.TotalsRowLabel))
                        tableColumn1.TotalsRowLabel = xlField.TotalsRowLabel;
                }
                tableColumns1.AppendChild(tableColumn1);
            }

            var tableStyleInfo1 = new TableStyleInfo
                                      {
                                          Name = Enum.GetName(typeof(XLTableTheme), xlTable.Theme),
                                          ShowFirstColumn = xlTable.EmphasizeFirstColumn,
                                          ShowLastColumn = xlTable.EmphasizeLastColumn,
                                          ShowRowStripes = xlTable.ShowRowStripes,
                                          ShowColumnStripes = xlTable.ShowColumnStripes
                                      };

            if (xlTable.ShowAutoFilter)
            {
                var autoFilter1 = new AutoFilter();

                if (xlTable.ShowTotalsRow)
                {
                    autoFilter1.Reference = xlTable.RangeAddress.FirstAddress + ":" +
                                            ExcelHelper.GetColumnLetterFromNumber(
                                                xlTable.RangeAddress.LastAddress.ColumnNumber) +
                                            (xlTable.RangeAddress.LastAddress.RowNumber - 1).ToStringLookup();
                }
                else
                    autoFilter1.Reference = reference;

                table.AppendChild(autoFilter1);
            }

            table.AppendChild(tableColumns1);
            table.AppendChild(tableStyleInfo1);

            tableDefinitionPart.Table = table;
        }

        #region GenerateWorkbookStylesPartContent

        private void GenerateWorkbookStylesPartContent(WorkbookStylesPart workbookStylesPart, SaveContext context)
        {
            var defaultStyle = new XLStyle(null, DefaultStyle);
            Int32 defaultStyleId = GetStyleId(defaultStyle);
            if (!context.SharedFonts.ContainsKey(defaultStyle.Font))
                context.SharedFonts.Add(defaultStyle.Font, new FontInfo {FontId = 0, Font = defaultStyle.Font});

            var sharedFills = new Dictionary<IXLFill, FillInfo>
                                  {{defaultStyle.Fill, new FillInfo {FillId = 2, Fill = defaultStyle.Fill}}};

            var sharedBorders = new Dictionary<IXLBorder, BorderInfo>
                                    {{defaultStyle.Border, new BorderInfo {BorderId = 0, Border = defaultStyle.Border}}};

            var sharedNumberFormats = new Dictionary<IXLNumberFormat, NumberFormatInfo>
                                          {
                                              {
                                                  defaultStyle.NumberFormat,
                                                  new NumberFormatInfo
                                                      {NumberFormatId = 0, NumberFormat = defaultStyle.NumberFormat}
                                                  }
                                          };

            //Dictionary<String, AlignmentInfo> sharedAlignments = new Dictionary<String, AlignmentInfo>();
            //sharedAlignments.Add(defaultStyle.Alignment.ToString(), new AlignmentInfo() { AlignmentId = 0, Alignment = defaultStyle.Alignment });

            if (workbookStylesPart.Stylesheet == null)
                workbookStylesPart.Stylesheet = new Stylesheet();

            // Cell styles = Named styles
            if (workbookStylesPart.Stylesheet.CellStyles == null)
                workbookStylesPart.Stylesheet.CellStyles = new CellStyles();

            UInt32 defaultFormatId;
            if (workbookStylesPart.Stylesheet.CellStyles.Elements<CellStyle>().Any(c => c.Name == "Normal"))
            {
                defaultFormatId =
                    workbookStylesPart.Stylesheet.CellStyles.Elements<CellStyle>().Where(c => c.Name == "Normal").Single
                        ().FormatId.Value;
            }
            else if (workbookStylesPart.Stylesheet.CellStyles.Elements<CellStyle>().Any())
            {
                defaultFormatId =
                    workbookStylesPart.Stylesheet.CellStyles.Elements<CellStyle>().Max(c => c.FormatId.Value) + 1;
            }
            else
                defaultFormatId = 0;

            context.SharedStyles.Add(defaultStyleId,
                                     new StyleInfo
                                         {
                                             StyleId = defaultFormatId,
                                             Style = defaultStyle,
                                             FontId = 0,
                                             FillId = 0,
                                             BorderId = 0,
                                             NumberFormatId = 0
                                             //AlignmentId = 0
                                         });

            UInt32 styleCount = 1;
            UInt32 fontCount = 1;
            UInt32 fillCount = 3;
            UInt32 borderCount = 1;
            Int32 numberFormatCount = 1;
            var xlStyles = new HashSet<Int32>();

            foreach (XLWorksheet worksheet in WorksheetsInternal)
            {
                
                foreach (var s in worksheet.GetStyleIds().Where(s => !xlStyles.Contains(s)))
                    xlStyles.Add(s);

                foreach (
                    Int32 s in
                        worksheet.Internals.ColumnsCollection.Select(kp => kp.Value.GetStyleId()).Where(
                            s => !xlStyles.Contains(s)))
                    xlStyles.Add(s);

                foreach (
                    Int32 s in
                        worksheet.Internals.RowsCollection.Select(kp => kp.Value.GetStyleId()).Where(s => !xlStyles.Contains(s))
                    )
                    xlStyles.Add(s);
            }

            foreach (var xlStyle in xlStyles.Select(GetStyleById))
            {
                if (!context.SharedFonts.ContainsKey(xlStyle.Font))
                    context.SharedFonts.Add(xlStyle.Font, new FontInfo {FontId = fontCount++, Font = xlStyle.Font});

                if (!sharedFills.ContainsKey(xlStyle.Fill))
                    sharedFills.Add(xlStyle.Fill, new FillInfo {FillId = fillCount++, Fill = xlStyle.Fill});

                if (!sharedBorders.ContainsKey(xlStyle.Border))
                    sharedBorders.Add(xlStyle.Border, new BorderInfo {BorderId = borderCount++, Border = xlStyle.Border});

                if (   xlStyle.NumberFormat.NumberFormatId != -1 
                       || sharedNumberFormats.ContainsKey(xlStyle.NumberFormat))
                    continue;

                sharedNumberFormats.Add(xlStyle.NumberFormat,
                                        new NumberFormatInfo
                                            {
                                                NumberFormatId = numberFormatCount + 164,
                                                NumberFormat = xlStyle.NumberFormat
                                            });
                numberFormatCount++;
            }

            var allSharedNumberFormats = ResolveNumberFormats(workbookStylesPart, sharedNumberFormats);
            ResolveFonts(workbookStylesPart, context);
            var allSharedFills = ResolveFills(workbookStylesPart, sharedFills);
            var allSharedBorders = ResolveBorders(workbookStylesPart, sharedBorders);

            foreach (Int32 id in xlStyles)
            {
                var xlStyle = GetStyleById(id);
                if (context.SharedStyles.ContainsKey(id)) continue;

                int numberFormatId = xlStyle.NumberFormat.NumberFormatId >= 0
                                         ? xlStyle.NumberFormat.NumberFormatId
                                         : allSharedNumberFormats[xlStyle.NumberFormat].NumberFormatId;

                context.SharedStyles.Add(id,
                                         new StyleInfo
                                             {
                                                 StyleId = styleCount++,
                                                 Style = xlStyle,
                                                 FontId = context.SharedFonts[xlStyle.Font].FontId,
                                                 FillId = allSharedFills[xlStyle.Fill].FillId,
                                                 BorderId = allSharedBorders[xlStyle.Border].BorderId,
                                                 NumberFormatId = numberFormatId
                                             });
            }

            ResolveCellStyleFormats(workbookStylesPart, context);
            ResolveRest(workbookStylesPart, context);

            if (!workbookStylesPart.Stylesheet.CellStyles.Elements<CellStyle>().Any(c => c.Name == "Normal"))
            {
                //var defaultFormatId = context.SharedStyles.Values.Where(s => s.Style.Equals(DefaultStyle)).Single().StyleId;

                var cellStyle1 = new CellStyle {Name = "Normal", FormatId = defaultFormatId, BuiltinId = 0U};
                workbookStylesPart.Stylesheet.CellStyles.AppendChild(cellStyle1);
            }
            workbookStylesPart.Stylesheet.CellStyles.Count = (UInt32)workbookStylesPart.Stylesheet.CellStyles.Count();

            var newSharedStyles = new Dictionary<Int32, StyleInfo>();
            foreach (KeyValuePair<Int32, StyleInfo> ss in context.SharedStyles)
            {
                Int32 styleId = -1;
                foreach (CellFormat f in workbookStylesPart.Stylesheet.CellFormats)
                {
                    styleId++;
                    if (CellFormatsAreEqual(f, ss.Value))
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

            //TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium9", DefaultPivotStyle = "PivotStyleLight16" };
            //workbookStylesPart.Stylesheet.AppendChild(tableStyles1);
        }

        private static void ResolveRest(WorkbookStylesPart workbookStylesPart, SaveContext context)
        {
            if (workbookStylesPart.Stylesheet.CellFormats == null)
                workbookStylesPart.Stylesheet.CellFormats = new CellFormats();

            foreach (StyleInfo styleInfo in context.SharedStyles.Values)
            {
                var info = styleInfo;
                Boolean foundOne =
                    workbookStylesPart.Stylesheet.CellFormats.Cast<CellFormat>().Any(f => CellFormatsAreEqual(f, info));
                
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
                                        TextRotation = (UInt32)styleInfo.Style.Alignment.TextRotation,
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
        }

        private static void ResolveCellStyleFormats(WorkbookStylesPart workbookStylesPart,
                                                    SaveContext context)
        {
            if (workbookStylesPart.Stylesheet.CellStyleFormats == null)
                workbookStylesPart.Stylesheet.CellStyleFormats = new CellStyleFormats();

            foreach (StyleInfo styleInfo in context.SharedStyles.Values)
            {
                var info = styleInfo;
                Boolean foundOne =
                    workbookStylesPart.Stylesheet.CellStyleFormats.Cast<CellFormat>().Any(
                        f => CellFormatsAreEqual(f, info));

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
                                     ApplyNumberFormat = false,
                                     ApplyFill = ApplyFill(styleInfo),
                                     ApplyBorder = ApplyBorder(styleInfo),
                                     ApplyAlignment = false,
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

        private static bool CellFormatsAreEqual(CellFormat f, StyleInfo styleInfo)
        {
            return
                styleInfo.BorderId == f.BorderId
                && styleInfo.FillId == f.FillId
                && styleInfo.FontId == f.FontId
                && styleInfo.NumberFormatId == f.NumberFormatId
                && f.ApplyNumberFormat != null && f.ApplyNumberFormat == false
                && f.ApplyAlignment != null && f.ApplyAlignment == false
                && f.ApplyFill != null && f.ApplyFill == ApplyFill(styleInfo)
                && f.ApplyBorder != null && f.ApplyBorder == ApplyBorder(styleInfo)
                && AlignmentsAreEqual(f.Alignment, styleInfo.Style.Alignment)
                && ProtectionsAreEqual(f.Protection, styleInfo.Style.Protection)
                ;
        }

        private static bool ProtectionsAreEqual(Protection protection, IXLProtection xlProtection)
        {
            var p = new XLProtection();
            if (protection != null)
            {
                if (protection.Locked != null)
                    p.Locked = protection.Locked.Value;
                if (protection.Hidden != null)
                    p.Hidden = protection.Hidden.Value;
            }
            return p.Equals(xlProtection);
        }

        private static bool AlignmentsAreEqual(Alignment alignment, IXLAlignment xlAlignment)
        {
            var a = new XLAlignment();
            if (alignment != null)
            {
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
                    a.TextRotation = (Int32)alignment.TextRotation.Value;
                if (alignment.ShrinkToFit != null)
                    a.ShrinkToFit = alignment.ShrinkToFit.Value;
                if (alignment.RelativeIndent != null)
                    a.RelativeIndent = alignment.RelativeIndent.Value;
                if (alignment.JustifyLastLine != null)
                    a.JustifyLastLine = alignment.JustifyLastLine.Value;
            }
            return a.Equals(xlAlignment);
        }

        private Dictionary<IXLBorder, BorderInfo> ResolveBorders(WorkbookStylesPart workbookStylesPart,
                                                                 Dictionary<IXLBorder, BorderInfo> sharedBorders)
        {
            if (workbookStylesPart.Stylesheet.Borders == null)
                workbookStylesPart.Stylesheet.Borders = new Borders();

            var allSharedBorders = new Dictionary<IXLBorder, BorderInfo>();
            foreach (BorderInfo borderInfo in sharedBorders.Values)
            {
                Int32 borderId = 0;
                Boolean foundOne = false;
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
                                     new BorderInfo {Border = borderInfo.Border, BorderId = (UInt32)borderId});
            }
            workbookStylesPart.Stylesheet.Borders.Count = (UInt32)workbookStylesPart.Stylesheet.Borders.Count();
            return allSharedBorders;
        }

        private static Border GetNewBorder(BorderInfo borderInfo)
        {
            var border = new Border
                             {DiagonalUp = borderInfo.Border.DiagonalUp, DiagonalDown = borderInfo.Border.DiagonalDown};

            var leftBorder = new LeftBorder {Style = borderInfo.Border.LeftBorder.ToOpenXml()};
            var leftBorderColor = GetNewColor(borderInfo.Border.LeftBorderColor);
            leftBorder.AppendChild(leftBorderColor);
            border.AppendChild(leftBorder);

            var rightBorder = new RightBorder {Style = borderInfo.Border.RightBorder.ToOpenXml()};
            var rightBorderColor = GetNewColor(borderInfo.Border.RightBorderColor);
            rightBorder.AppendChild(rightBorderColor);
            border.AppendChild(rightBorder);

            var topBorder = new TopBorder {Style = borderInfo.Border.TopBorder.ToOpenXml()};
            var topBorderColor = GetNewColor(borderInfo.Border.TopBorderColor);
            topBorder.AppendChild(topBorderColor);
            border.AppendChild(topBorder);

            var bottomBorder = new BottomBorder {Style = borderInfo.Border.BottomBorder.ToOpenXml()};
            var bottomBorderColor = GetNewColor(borderInfo.Border.BottomBorderColor);
            bottomBorder.AppendChild(bottomBorderColor);
            border.AppendChild(bottomBorder);

            var diagonalBorder = new DiagonalBorder {Style = borderInfo.Border.DiagonalBorder.ToOpenXml()};
            var diagonalBorderColor = GetNewColor(borderInfo.Border.DiagonalBorderColor);
            diagonalBorder.AppendChild(diagonalBorderColor);
            border.AppendChild(diagonalBorder);

            return border;
        }

        private bool BordersAreEqual(Border b, IXLBorder xlBorder)
        {
            var nb = new XLBorder();
            if (b.DiagonalUp != null)
                nb.DiagonalUp = b.DiagonalUp.Value;

            if (b.DiagonalDown != null)
                nb.DiagonalDown = b.DiagonalDown.Value;

            if (b.LeftBorder != null)
            {
                if (b.LeftBorder.Style != null)
                    nb.LeftBorder = b.LeftBorder.Style.Value.ToClosedXml();
                var bColor = GetColor(b.LeftBorder.Color);
                if (bColor.HasValue)
                    nb.LeftBorderColor = bColor;
            }

            if (b.RightBorder != null)
            {
                if (b.RightBorder.Style != null)
                    nb.RightBorder = b.RightBorder.Style.Value.ToClosedXml();
                var bColor = GetColor(b.RightBorder.Color);
                if (bColor.HasValue)
                    nb.RightBorderColor = bColor;
            }

            if (b.TopBorder != null)
            {
                if (b.TopBorder.Style != null)
                    nb.TopBorder = b.TopBorder.Style.Value.ToClosedXml();
                var bColor = GetColor(b.TopBorder.Color);
                if (bColor.HasValue)
                    nb.TopBorderColor = bColor;
            }

            if (b.BottomBorder != null)
            {
                if (b.BottomBorder.Style != null)
                    nb.BottomBorder = b.BottomBorder.Style.Value.ToClosedXml();
                var bColor = GetColor(b.BottomBorder.Color);
                if (bColor.HasValue)
                    nb.BottomBorderColor = bColor;
            }

            return nb.Equals(xlBorder);
        }

        private Dictionary<IXLFill, FillInfo> ResolveFills(WorkbookStylesPart workbookStylesPart,
                                                           Dictionary<IXLFill, FillInfo> sharedFills)
        {
            if (workbookStylesPart.Stylesheet.Fills == null)
                workbookStylesPart.Stylesheet.Fills = new Fills();

            ResolveFillWithPattern(workbookStylesPart.Stylesheet.Fills, PatternValues.None);
            ResolveFillWithPattern(workbookStylesPart.Stylesheet.Fills, PatternValues.Gray125);

            var allSharedFills = new Dictionary<IXLFill, FillInfo>();
            foreach (FillInfo fillInfo in sharedFills.Values)
            {
                Int32 fillId = 0;
                Boolean foundOne = false;
                foreach (Fill f in workbookStylesPart.Stylesheet.Fills)
                {
                    if (FillsAreEqual(f, fillInfo.Fill))
                    {
                        foundOne = true;
                        break;
                    }
                    fillId++;
                }
                if (!foundOne)
                {
                    var fill = GetNewFill(fillInfo);
                    workbookStylesPart.Stylesheet.Fills.AppendChild(fill);
                }
                allSharedFills.Add(fillInfo.Fill, new FillInfo {Fill = fillInfo.Fill, FillId = (UInt32)fillId});
            }

            workbookStylesPart.Stylesheet.Fills.Count = (UInt32)workbookStylesPart.Stylesheet.Fills.Count();
            return allSharedFills;
        }

        private static void ResolveFillWithPattern(Fills fills, PatternValues patternValues)
        {
            if (fills.Elements<Fill>().Any(f =>
                                           f.PatternFill.PatternType == patternValues
                                           && f.PatternFill.ForegroundColor == null
                                           && f.PatternFill.BackgroundColor == null
                )) return;

            var fill1 = new Fill();
            var patternFill1 = new PatternFill {PatternType = patternValues};
            fill1.AppendChild(patternFill1);
            fills.AppendChild(fill1);
        }

        private static Fill GetNewFill(FillInfo fillInfo)
        {
            var fill = new Fill();

            var patternFill = new PatternFill {PatternType = fillInfo.Fill.PatternType.ToOpenXml()};
            var foregroundColor = new ForegroundColor();
            if (fillInfo.Fill.PatternColor.ColorType == XLColorType.Color)
                foregroundColor.Rgb = fillInfo.Fill.PatternColor.Color.ToHex();
            else if (fillInfo.Fill.PatternColor.ColorType == XLColorType.Indexed)
                foregroundColor.Indexed = (UInt32)fillInfo.Fill.PatternColor.Indexed;
            else
            {
                foregroundColor.Theme = (UInt32)fillInfo.Fill.PatternColor.ThemeColor;
                if (fillInfo.Fill.PatternColor.ThemeTint != 1)
                    foregroundColor.Tint = fillInfo.Fill.PatternColor.ThemeTint;
            }
            var backgroundColor = new BackgroundColor();
            if (fillInfo.Fill.PatternBackgroundColor.ColorType == XLColorType.Color)
                backgroundColor.Rgb = fillInfo.Fill.PatternBackgroundColor.Color.ToHex();
            else if (fillInfo.Fill.PatternBackgroundColor.ColorType == XLColorType.Indexed)
                backgroundColor.Indexed = (UInt32)fillInfo.Fill.PatternBackgroundColor.Indexed;
            else
            {
                backgroundColor.Theme = (UInt32)fillInfo.Fill.PatternBackgroundColor.ThemeColor;
                if (fillInfo.Fill.PatternBackgroundColor.ThemeTint != 1)
                    backgroundColor.Tint = fillInfo.Fill.PatternBackgroundColor.ThemeTint;
            }

            patternFill.AppendChild(foregroundColor);
            patternFill.AppendChild(backgroundColor);

            fill.AppendChild(patternFill);

            return fill;
        }

        private bool FillsAreEqual(Fill f, IXLFill xlFill)
        {
            var nF = new XLFill();
            if (f.PatternFill != null)
            {
                if (f.PatternFill.PatternType != null)
                    nF.PatternType = f.PatternFill.PatternType.Value.ToClosedXml();

                var fColor = GetColor(f.PatternFill.ForegroundColor);
                if (fColor.HasValue)
                    nF.PatternColor = fColor;

                var bColor = GetColor(f.PatternFill.BackgroundColor);
                if (bColor.HasValue)
                    nF.PatternBackgroundColor = bColor;
            }
            return nF.Equals(xlFill);
        }

        private void ResolveFonts(WorkbookStylesPart workbookStylesPart, SaveContext context)
        {
            if (workbookStylesPart.Stylesheet.Fonts == null)
                workbookStylesPart.Stylesheet.Fonts = new Fonts();

            var newFonts = new Dictionary<IXLFont, FontInfo>();
            foreach (FontInfo fontInfo in context.SharedFonts.Values)
            {
                Int32 fontId = 0;
                Boolean foundOne = false;
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
                newFonts.Add(fontInfo.Font, new FontInfo {Font = fontInfo.Font, FontId = (UInt32)fontId});
            }
            context.SharedFonts.Clear();
            foreach (KeyValuePair<IXLFont, FontInfo> kp in newFonts)
                context.SharedFonts.Add(kp.Key, kp.Value);

            workbookStylesPart.Stylesheet.Fonts.Count = (UInt32)workbookStylesPart.Stylesheet.Fonts.Count();
        }

        private static Font GetNewFont(FontInfo fontInfo)
        {
            var font = new Font();
            var bold = fontInfo.Font.Bold ? new Bold() : null;
            var italic = fontInfo.Font.Italic ? new Italic() : null;
            var underline = fontInfo.Font.Underline != XLFontUnderlineValues.None
                                ? new Underline {Val = fontInfo.Font.Underline.ToOpenXml()}
                                : null;
            var strike = fontInfo.Font.Strikethrough ? new Strike() : null;
            var verticalAlignment = new VerticalTextAlignment {Val = fontInfo.Font.VerticalAlignment.ToOpenXml()};
            var shadow = fontInfo.Font.Shadow ? new Shadow() : null;
            var fontSize = new FontSize {Val = fontInfo.Font.FontSize};
            var color = GetNewColor(fontInfo.Font.FontColor);

            var fontName = new FontName {Val = fontInfo.Font.FontName};
            var fontFamilyNumbering = new FontFamilyNumbering {Val = (Int32)fontInfo.Font.FontFamilyNumbering};

            if (bold != null)
                font.AppendChild(bold);
            if (italic != null)
                font.AppendChild(italic);
            if (underline != null)
                font.AppendChild(underline);
            if (strike != null)
                font.AppendChild(strike);
            font.AppendChild(verticalAlignment);
            if (shadow != null)
                font.AppendChild(shadow);
            font.AppendChild(fontSize);
            font.AppendChild(color);
            font.AppendChild(fontName);
            font.AppendChild(fontFamilyNumbering);

            return font;
        }

        private static Color GetNewColor(IXLColor xlColor)
        {
            var color = new Color();
            if (xlColor.ColorType == XLColorType.Color)
                color.Rgb = xlColor.Color.ToHex();
            else if (xlColor.ColorType == XLColorType.Indexed)
                color.Indexed = (UInt32)xlColor.Indexed;
            else
            {
                color.Theme = (UInt32)xlColor.ThemeColor;
                if (xlColor.ThemeTint != 1)
                    color.Tint = xlColor.ThemeTint;
            }
            return color;
        }

        private static TabColor GetTabColor(IXLColor xlColor)
        {
            var color = new TabColor();
            if (xlColor.ColorType == XLColorType.Color)
                color.Rgb = xlColor.Color.ToHex();
            else if (xlColor.ColorType == XLColorType.Indexed)
                color.Indexed = (UInt32)xlColor.Indexed;
            else
            {
                color.Theme = (UInt32)xlColor.ThemeColor;
                if (xlColor.ThemeTint != 1)
                    color.Tint = xlColor.ThemeTint;
            }
            return color;
        }

        private bool FontsAreEqual(Font f, IXLFont xlFont)
        {
            var nf = new XLFont {Bold = f.Bold != null, Italic = f.Italic != null};
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
            var fColor = GetColor(f.Color);
            if (fColor.HasValue)
                nf.FontColor = fColor;
            if (f.FontName != null)
                nf.FontName = f.FontName.Val;
            if (f.FontFamilyNumbering != null)
                nf.FontFamilyNumbering = (XLFontFamilyNumberingValues)f.FontFamilyNumbering.Val.Value;

            return nf.Equals(xlFont);
        }

        private static Dictionary<IXLNumberFormat, NumberFormatInfo> ResolveNumberFormats(
            WorkbookStylesPart workbookStylesPart,
            Dictionary<IXLNumberFormat, NumberFormatInfo> sharedNumberFormats)
        {
            if (workbookStylesPart.Stylesheet.NumberingFormats == null)
                workbookStylesPart.Stylesheet.NumberingFormats = new NumberingFormats();

            var allSharedNumberFormats = new Dictionary<IXLNumberFormat, NumberFormatInfo>();
            foreach (NumberFormatInfo numberFormatInfo in sharedNumberFormats.Values)
            {
                Int32 numberingFormatId = 0;
                Boolean foundOne = false;
                foreach (NumberingFormat nf in workbookStylesPart.Stylesheet.NumberingFormats)
                {
                    if (NumberFormatsAreEqual(nf, numberFormatInfo.NumberFormat))
                    {
                        foundOne = true;
                        numberingFormatId = (Int32)nf.NumberFormatId.Value;
                        break;
                    }
                    numberingFormatId++;
                }
                if (!foundOne)
                {
                    var numberingFormat = new NumberingFormat
                                              {
                                                  NumberFormatId = (UInt32)numberingFormatId,
                                                  FormatCode = numberFormatInfo.NumberFormat.Format
                                              };
                    workbookStylesPart.Stylesheet.NumberingFormats.AppendChild(numberingFormat);
                }
                allSharedNumberFormats.Add(numberFormatInfo.NumberFormat,
                                           new NumberFormatInfo
                                               {
                                                   NumberFormat = numberFormatInfo.NumberFormat,
                                                   NumberFormatId = numberingFormatId
                                               });
            }
            workbookStylesPart.Stylesheet.NumberingFormats.Count =
                (UInt32)workbookStylesPart.Stylesheet.NumberingFormats.Count();
            return allSharedNumberFormats;
        }

        private static bool NumberFormatsAreEqual(NumberingFormat nf, IXLNumberFormat xlNumberFormat)
        {
            var newXLNumberFormat = new XLNumberFormat();

            if (nf.FormatCode != null && !StringExtensions.IsNullOrWhiteSpace(nf.FormatCode.Value))
                newXLNumberFormat.Format = nf.FormatCode.Value;
            else if (nf.NumberFormatId != null)
                newXLNumberFormat.NumberFormatId = (Int32)nf.NumberFormatId.Value;

            return newXLNumberFormat.Equals(xlNumberFormat);
        }

        #endregion

        #region GenerateWorksheetPartContent

        private static void GenerateWorksheetPartContent(WorksheetPart worksheetPart, XLWorksheet xlWorksheet,
                                                         SaveContext context)
        {
            #region Worksheet

            if (worksheetPart.Worksheet == null)
                worksheetPart.Worksheet = new Worksheet();

            GenerateTables(xlWorksheet, worksheetPart, context);

            if (
                !worksheetPart.Worksheet.NamespaceDeclarations.Contains(new KeyValuePair<String, String>("r",
                                                                                                         "http://schemas.openxmlformats.org/officeDocument/2006/relationships")))
            {
                worksheetPart.Worksheet.AddNamespaceDeclaration("r",
                                                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            }

            #endregion

            var cm = new XLWSContentManager(worksheetPart.Worksheet);

            #region SheetProperties

            if (worksheetPart.Worksheet.SheetProperties == null)
                worksheetPart.Worksheet.SheetProperties = new SheetProperties();

            worksheetPart.Worksheet.SheetProperties.TabColor = xlWorksheet.TabColor.HasValue
                                                                   ? GetTabColor(xlWorksheet.TabColor)
                                                                   : null;

            cm.SetElement(XLWSContentManager.XLWSContents.SheetProperties, worksheetPart.Worksheet.SheetProperties);

            if (worksheetPart.Worksheet.SheetProperties.OutlineProperties == null)
                worksheetPart.Worksheet.SheetProperties.OutlineProperties = new OutlineProperties();

            worksheetPart.Worksheet.SheetProperties.OutlineProperties.SummaryBelow =
                (xlWorksheet.Outline.SummaryVLocation ==
                 XLOutlineSummaryVLocation.Bottom);
            worksheetPart.Worksheet.SheetProperties.OutlineProperties.SummaryRight =
                (xlWorksheet.Outline.SummaryHLocation ==
                 XLOutlineSummaryHLocation.Right);

            if (worksheetPart.Worksheet.SheetProperties.PageSetupProperties == null
                && (xlWorksheet.PageSetup.PagesTall > 0 || xlWorksheet.PageSetup.PagesWide > 0))
                worksheetPart.Worksheet.SheetProperties.PageSetupProperties = new PageSetupProperties { FitToPage = true };

            #endregion

            Int32 maxColumn = 0;

            String sheetDimensionReference = "A1";
            if (xlWorksheet.Internals.CellsCollection.Count > 0)
            {
                maxColumn = xlWorksheet.Internals.CellsCollection.MaxColumnUsed;
                Int32 maxRow = xlWorksheet.Internals.CellsCollection.MaxRowUsed;
                sheetDimensionReference = "A1:" + ExcelHelper.GetColumnLetterFromNumber(maxColumn) +
                                          maxRow.ToStringLookup();
            }

            if (xlWorksheet.Internals.ColumnsCollection.Count > 0)
            {
                Int32 maxColCollection = xlWorksheet.Internals.ColumnsCollection.Keys.Max();
                if (maxColCollection > maxColumn)
                    maxColumn = maxColCollection;
            }

            #region SheetViews

            if (worksheetPart.Worksheet.SheetDimension == null)
                worksheetPart.Worksheet.SheetDimension = new SheetDimension { Reference = sheetDimensionReference };

            cm.SetElement(XLWSContentManager.XLWSContents.SheetDimension, worksheetPart.Worksheet.SheetDimension);

            if (worksheetPart.Worksheet.SheetViews == null)
                worksheetPart.Worksheet.SheetViews = new SheetViews();

            cm.SetElement(XLWSContentManager.XLWSContents.SheetViews, worksheetPart.Worksheet.SheetViews);

            var sheetView = (SheetView)worksheetPart.Worksheet.SheetViews.FirstOrDefault();
            if (sheetView == null)
            {
                sheetView = new SheetView { WorkbookViewId = 0U };
                worksheetPart.Worksheet.SheetViews.AppendChild(sheetView);
            }

            sheetView.TabSelected = xlWorksheet.TabSelected;

            if (xlWorksheet.ShowFormulas)
                sheetView.ShowFormulas = true;
            else
                sheetView.ShowFormulas = null;

            if (xlWorksheet.ShowGridLines)
                sheetView.ShowGridLines = null;
            else
                sheetView.ShowGridLines = false;

            if (xlWorksheet.ShowOutlineSymbols)
                sheetView.ShowOutlineSymbols = null;
            else
                sheetView.ShowOutlineSymbols = false;

            if (xlWorksheet.ShowRowColHeaders)
                sheetView.ShowRowColHeaders = null;
            else
                sheetView.ShowRowColHeaders = false;

            if (xlWorksheet.ShowRuler)
                sheetView.ShowRuler = null;
            else
                sheetView.ShowRuler = false;

            if (xlWorksheet.ShowWhiteSpace)
                sheetView.ShowWhiteSpace = null;
            else
                sheetView.ShowWhiteSpace = false;

            if (xlWorksheet.ShowZeros)
                sheetView.ShowZeros = null;
            else
                sheetView.ShowZeros = false;

            if (xlWorksheet.RightToLeft)
                sheetView.RightToLeft = true;
            else
                sheetView.RightToLeft = null;

            var pane = sheetView.Elements<Pane>().FirstOrDefault();
            if (pane == null)
            {
                pane = new Pane();
                sheetView.AppendChild(pane);
            }

            pane.State = PaneStateValues.FrozenSplit;
            Double hSplit = xlWorksheet.SheetView.SplitColumn;
            Double ySplit = xlWorksheet.SheetView.SplitRow;


            pane.HorizontalSplit = hSplit;
            pane.VerticalSplit = ySplit;

            pane.TopLeftCell = ExcelHelper.GetColumnLetterFromNumber(xlWorksheet.SheetView.SplitColumn + 1)
                               + (xlWorksheet.SheetView.SplitRow + 1);

            if (hSplit == 0 && ySplit == 0)
                sheetView.RemoveAllChildren<Pane>();

            #endregion

            int maxOutlineColumn = 0;
            if (xlWorksheet.ColumnCount() > 0)
                maxOutlineColumn = xlWorksheet.GetMaxColumnOutline();

            int maxOutlineRow = 0;
            if (xlWorksheet.RowCount() > 0)
                maxOutlineRow = xlWorksheet.GetMaxRowOutline();

            #region SheetFormatProperties

            if (worksheetPart.Worksheet.SheetFormatProperties == null)
                worksheetPart.Worksheet.SheetFormatProperties = new SheetFormatProperties();

            cm.SetElement(XLWSContentManager.XLWSContents.SheetFormatProperties,
                          worksheetPart.Worksheet.SheetFormatProperties);

            worksheetPart.Worksheet.SheetFormatProperties.DefaultRowHeight = xlWorksheet.RowHeight;

            if (xlWorksheet.RowHeightChanged)
                worksheetPart.Worksheet.SheetFormatProperties.CustomHeight = true;
            else
                worksheetPart.Worksheet.SheetFormatProperties.CustomHeight = null;


            double worksheetColumnWidth = GetColumnWidth(xlWorksheet.ColumnWidth);
            if (xlWorksheet.ColumnWidthChanged)
                worksheetPart.Worksheet.SheetFormatProperties.DefaultColumnWidth = worksheetColumnWidth;
            else
                worksheetPart.Worksheet.SheetFormatProperties.DefaultColumnWidth = null;


            if (maxOutlineColumn > 0)
                worksheetPart.Worksheet.SheetFormatProperties.OutlineLevelColumn = (byte)maxOutlineColumn;
            else
                worksheetPart.Worksheet.SheetFormatProperties.OutlineLevelColumn = null;

            if (maxOutlineRow > 0)
                worksheetPart.Worksheet.SheetFormatProperties.OutlineLevelRow = (byte)maxOutlineRow;
            else
                worksheetPart.Worksheet.SheetFormatProperties.OutlineLevelRow = null;

            #endregion

            #region Columns

            if (xlWorksheet.Internals.CellsCollection.Count == 0 &&
                xlWorksheet.Internals.ColumnsCollection.Count == 0
                && xlWorksheet.Style.Equals(DefaultStyle))
                worksheetPart.Worksheet.RemoveAllChildren<Columns>();
            else
            {
                if (!worksheetPart.Worksheet.Elements<Columns>().Any())
                {
                    var previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.Columns);
                    worksheetPart.Worksheet.InsertAfter(new Columns(), previousElement);
                }

                var columns = worksheetPart.Worksheet.Elements<Columns>().First();
                cm.SetElement(XLWSContentManager.XLWSContents.Columns, columns);

                var sheetColumnsByMin = columns.Elements<Column>().ToDictionary(c => c.Min.Value, c => c);
                //Dictionary<UInt32, Column> sheetColumnsByMax = columns.Elements<Column>().ToDictionary(c => c.Max.Value, c => c);

                Int32 minInColumnsCollection;
                Int32 maxInColumnsCollection;
                if (xlWorksheet.Internals.ColumnsCollection.Count > 0)
                {
                    minInColumnsCollection = xlWorksheet.Internals.ColumnsCollection.Keys.Min();
                    maxInColumnsCollection = xlWorksheet.Internals.ColumnsCollection.Keys.Max();
                }
                else
                {
                    minInColumnsCollection = 1;
                    maxInColumnsCollection = 0;
                }

                uint worksheetStyleId = context.SharedStyles[xlWorksheet.GetStyleId()].StyleId;
                if (minInColumnsCollection > 1)
                {
                    UInt32Value min = 1;
                    UInt32Value max = (UInt32)(minInColumnsCollection - 1);


                    for (var co = min; co <= max; co++)
                    {
                        var column = new Column
                                         {
                                             Min = co,
                                             Max = co,
                                             Style = worksheetStyleId,
                                             Width = worksheetColumnWidth,
                                             CustomWidth = true
                                         };

                        UpdateColumn(column, columns, sheetColumnsByMin); //, sheetColumnsByMax);
                    }
                }

                for (int co = minInColumnsCollection; co <= maxInColumnsCollection; co++)
                {
                    UInt32 styleId;
                    Double columnWidth;
                    Boolean isHidden = false;
                    Boolean collapsed = false;
                    Int32 outlineLevel = 0;
                    if (xlWorksheet.Internals.ColumnsCollection.ContainsKey(co))
                    {
                        styleId = context.SharedStyles[xlWorksheet.Internals.ColumnsCollection[co].GetStyleId()].StyleId;
                        columnWidth = GetColumnWidth(xlWorksheet.Internals.ColumnsCollection[co].Width);
                        isHidden = xlWorksheet.Internals.ColumnsCollection[co].IsHidden;
                        collapsed = xlWorksheet.Internals.ColumnsCollection[co].Collapsed;
                        outlineLevel = xlWorksheet.Internals.ColumnsCollection[co].OutlineLevel;
                    }
                    else
                    {
                        styleId = context.SharedStyles[xlWorksheet.GetStyleId()].StyleId;
                        columnWidth = worksheetColumnWidth;
                    }

                    var column = new Column
                                     {
                                         Min = (UInt32)co,
                                         Max = (UInt32)co,
                                         Style = styleId,
                                         Width = columnWidth,
                                         CustomWidth = true
                                     };

                    if (isHidden)
                        column.Hidden = true;
                    if (collapsed)
                        column.Collapsed = true;
                    if (outlineLevel > 0)
                        column.OutlineLevel = (byte)outlineLevel;

                    UpdateColumn(column, columns, sheetColumnsByMin); //, sheetColumnsByMax);
                }

                int collection = maxInColumnsCollection;
                foreach (
                    Column col in
                        columns.Elements<Column>().Where(c => c.Min > (UInt32)(collection)).OrderBy(
                            c => c.Min.Value))
                {
                    col.Style = worksheetStyleId;
                    col.Width = worksheetColumnWidth;
                    col.CustomWidth = true;

                    if ((Int32)col.Max.Value > maxInColumnsCollection)
                        maxInColumnsCollection = (Int32)col.Max.Value;
                }

                if (maxInColumnsCollection < ExcelHelper.MaxColumnNumber && !xlWorksheet.Style.Equals(DefaultStyle))
                {
                    var column = new Column
                                     {
                                         Min = (UInt32)(maxInColumnsCollection + 1),
                                         Max = (UInt32)(ExcelHelper.MaxColumnNumber),
                                         Style = worksheetStyleId,
                                         Width = worksheetColumnWidth,
                                         CustomWidth = true
                                     };
                    columns.AppendChild(column);
                }

                CollapseColumns(columns, sheetColumnsByMin);

                if (!columns.Any())
                {
                    worksheetPart.Worksheet.RemoveAllChildren<Columns>();
                    cm.SetElement(XLWSContentManager.XLWSContents.Columns, null);
                }
            }

            #endregion

            #region SheetData

            if (!worksheetPart.Worksheet.Elements<SheetData>().Any())
            {
                var previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.SheetData);
                worksheetPart.Worksheet.InsertAfter(new SheetData(), previousElement);
            }

            var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
            cm.SetElement(XLWSContentManager.XLWSContents.SheetData, sheetData);

            var cellsByRow = new Dictionary<Int32, List<IXLCell>>();
            foreach (XLCell c in xlWorksheet.Internals.CellsCollection.GetCells())
            {
                Int32 rowNum = c.Address.RowNumber;
                if (!cellsByRow.ContainsKey(rowNum))
                    cellsByRow.Add(rowNum, new List<IXLCell>());

                cellsByRow[rowNum].Add(c);
            }

            var sheetDataRows = sheetData.Elements<Row>().ToDictionary(r => (Int32)r.RowIndex.Value, r => r);
            foreach (
                KeyValuePair<int, XLRow> r in
                    xlWorksheet.Internals.RowsCollection.Deleted.Where(r => sheetDataRows.ContainsKey(r.Key)))
            {
                sheetData.RemoveChild(sheetDataRows[r.Key]);
                sheetDataRows.Remove(r.Key);
                xlWorksheet.Internals.CellsCollection.Deleted.RemoveWhere(d => d.Row == r.Key);
            }

            var distinctRows = cellsByRow.Keys.Union(xlWorksheet.Internals.RowsCollection.Keys);
            Boolean noRows = (sheetData.Elements<Row>().FirstOrDefault() == null);
            foreach (int distinctRow in distinctRows.OrderBy(r => r))
            {
                Row row; // = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex.Value == (UInt32)distinctRow);
                if (sheetDataRows.ContainsKey(distinctRow))
                    row = sheetDataRows[distinctRow];
                else
                {
                    row = new Row { RowIndex = (UInt32)distinctRow };
                    if (noRows)
                    {
                        sheetData.AppendChild(row);
                        noRows = false;
                    }
                    else
                    {
                        if (sheetDataRows.Any(r => r.Key > row.RowIndex.Value))
                        {
                            int minRow = sheetDataRows.Where(r => r.Key > (Int32)row.RowIndex.Value).Min(r => r.Key);
                            var rowBeforeInsert = sheetDataRows[minRow];
                            sheetData.InsertBefore(row, rowBeforeInsert);
                        }
                        else
                            sheetData.AppendChild(row);
                    }
                }

                if (maxColumn > 0)
                    row.Spans = new ListValue<StringValue> { InnerText = "1:" + maxColumn.ToStringLookup() };

                row.Height = null;
                row.CustomHeight = null;
                row.Hidden = null;
                row.StyleIndex = null;
                row.CustomFormat = null;
                row.Collapsed = null;
                if (xlWorksheet.Internals.RowsCollection.ContainsKey(distinctRow))
                {
                    var thisRow = xlWorksheet.Internals.RowsCollection[distinctRow];
                    if (thisRow.Height != xlWorksheet.RowHeight)
                    {
                        row.Height = thisRow.Height;
                        row.CustomHeight = true;
                    }

                    //if (!thisRow.Style.Equals(xlWorksheet.Style))
                    if (thisRow.GetStyleId() != xlWorksheet.GetStyleId())
                    {
                        row.StyleIndex = context.SharedStyles[thisRow.GetStyleId()].StyleId;
                        row.CustomFormat = true;
                    }
                    if (thisRow.IsHidden)
                        row.Hidden = true;
                    if (thisRow.Collapsed)
                        row.Collapsed = true;
                    if (thisRow.OutlineLevel > 0)
                        row.OutlineLevel = (byte)thisRow.OutlineLevel;
                }


                var cellsByReference = row.Elements<Cell>().ToDictionary(c => c.CellReference.Value, c => c);

                foreach (XLSheetPoint c in xlWorksheet.Internals.CellsCollection.Deleted.ToList())
                {
                    String key = ExcelHelper.GetColumnLetterFromNumber(c.Column) + c.Row.ToStringLookup();
                    if (!cellsByReference.ContainsKey(key)) continue;
                    row.RemoveChild(cellsByReference[key]);
                    xlWorksheet.Internals.CellsCollection.Deleted.Remove(c);
                }

                if (!cellsByRow.ContainsKey(distinctRow)) continue;

                Boolean isNewRow = !row.Elements<Cell>().Any();
                foreach (XLCell opCell in cellsByRow[distinctRow]
                    .OrderBy(c => c.Address.ColumnNumber)
                    .Select(c => (XLCell)c))
                {
                    uint styleId = context.SharedStyles[opCell.GetStyleId()].StyleId;

                    var dataType = opCell.DataType;
                    string cellReference = (opCell.Address).GetTrimmedAddress();

                    Cell cell;
                    if (cellsByReference.ContainsKey(cellReference))
                        cell = cellsByReference[cellReference];
                    else
                    {
                        cell = new Cell { CellReference = new StringValue(cellReference) };
                        if (isNewRow)
                            row.AppendChild(cell);
                        else
                        {
                            Int32 newColumn = ExcelHelper.GetColumnNumberFromAddress1(cellReference);

                            Cell cellBeforeInsert = null;
                            Int32 lastCo = Int32.MaxValue;
                            foreach (
                                Cell c in
                                    row.Elements<Cell>().Where(
                                        c =>
                                        ExcelHelper.GetColumnNumberFromAddress1(c.CellReference.Value) > newColumn))
                            {
                                int thidCo = ExcelHelper.GetColumnNumberFromAddress1(c.CellReference.Value);

                                if (lastCo <= thidCo) continue;

                                cellBeforeInsert = c;
                                lastCo = thidCo;
                            }
                            if (cellBeforeInsert == null)
                                row.AppendChild(cell);
                            else
                                row.InsertBefore(cell, cellBeforeInsert);
                        }
                    }

                    cell.StyleIndex = styleId;
                    if (!StringExtensions.IsNullOrWhiteSpace(opCell.FormulaA1))
                    {
                        String formula = opCell.FormulaA1;
                        if (formula.StartsWith("{"))
                        {
                            formula = formula.Substring(1, formula.Length - 2);
                            cell.CellFormula = new CellFormula(formula)
                                                   {
                                                       FormulaType = CellFormulaValues.Array,
                                                       Reference = cellReference
                                                   };
                        }
                        else
                            cell.CellFormula = new CellFormula(formula);
                        cell.CellValue = null;
                    }
                    else
                    {
                        cell.CellFormula = null;

                        cell.DataType = opCell.DataType == XLCellValues.DateTime ? null : GetCellValue(opCell);

                        var cellValue = new CellValue();
                        if (dataType == XLCellValues.Text)
                        {
                            if (opCell.InnerText.Length == 0)
                                cell.CellValue = null;
                            else
                            {
                                if (opCell.ShareString)
                                {
                                    cellValue.Text = opCell.SharedStringId.ToString();
                                    cell.CellValue = cellValue;
                                }
                                else
                                {
                                    String text = opCell.GetString();
                                    var t = new Text(text);
                                    if (text.PreserveSpaces())
                                        t.Space = SpaceProcessingModeValues.Preserve;

                                    cell.InlineString = new InlineString { Text = t };
                                }
                            }
                        }
                        else if (dataType == XLCellValues.TimeSpan)
                        {
                            var timeSpan = opCell.GetTimeSpan();
                            cellValue.Text =
                                XLCell.BaseDate.Add(timeSpan).ToOADate().ToString(CultureInfo.InvariantCulture);
                            cell.CellValue = cellValue;
                        }
                        else if (dataType == XLCellValues.DateTime || dataType == XLCellValues.Number)
                        {
                            cellValue.Text = Double.Parse(opCell.InnerText).ToString(CultureInfo.InvariantCulture);
                            cell.CellValue = cellValue;
                        }
                        else
                        {
                            cellValue.Text = opCell.InnerText;
                            cell.CellValue = cellValue;
                        }
                    }
                }
                xlWorksheet.Internals.CellsCollection.Deleted.RemoveWhere(d => d.Row == distinctRow);
            }
            foreach (var r in xlWorksheet.Internals.CellsCollection.Deleted.Select(c => c.Row).Distinct().Where(sheetDataRows.ContainsKey))
            {
                sheetData.RemoveChild(sheetDataRows[r]);
                sheetDataRows.Remove(r);
            }

            #endregion

            #region SheetProtection

            if (xlWorksheet.Protection.Protected)
            {
                if (!worksheetPart.Worksheet.Elements<SheetProtection>().Any())
                {
                    var previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.SheetProtection);
                    worksheetPart.Worksheet.InsertAfter(new SheetProtection(), previousElement);
                }

                var sheetProtection = worksheetPart.Worksheet.Elements<SheetProtection>().First();
                cm.SetElement(XLWSContentManager.XLWSContents.SheetProtection, sheetProtection);

                var protection = xlWorksheet.Protection;
                sheetProtection.Sheet = protection.Protected;
                if (!StringExtensions.IsNullOrWhiteSpace(protection.PasswordHash))
                    sheetProtection.Password = protection.PasswordHash;
                sheetProtection.FormatCells = GetBooleanValue(!protection.FormatCells, true);
                sheetProtection.FormatColumns = GetBooleanValue(!protection.FormatColumns, true);
                sheetProtection.FormatRows = GetBooleanValue(!protection.FormatRows, true);
                sheetProtection.InsertColumns = GetBooleanValue(!protection.InsertColumns, true);
                sheetProtection.InsertHyperlinks = GetBooleanValue(!protection.InsertHyperlinks, true);
                sheetProtection.InsertRows = GetBooleanValue(!protection.InsertRows, true);
                sheetProtection.DeleteColumns = GetBooleanValue(!protection.DeleteColumns, true);
                sheetProtection.DeleteRows = GetBooleanValue(!protection.DeleteRows, true);
                sheetProtection.AutoFilter = GetBooleanValue(!protection.AutoFilter, true);
                sheetProtection.PivotTables = GetBooleanValue(!protection.PivotTables, true);
                sheetProtection.Sort = GetBooleanValue(!protection.Sort, true);
                sheetProtection.SelectLockedCells = GetBooleanValue(!protection.SelectLockedCells, false);
                sheetProtection.SelectUnlockedCells = GetBooleanValue(!protection.SelectUnlockedCells, false);
            }
            else
            {
                worksheetPart.Worksheet.RemoveAllChildren<SheetProtection>();
                cm.SetElement(XLWSContentManager.XLWSContents.SheetProtection, null);
            }

            #endregion

            #region AutoFilter

            if (xlWorksheet.AutoFilterRange != null)
            {
                if (!worksheetPart.Worksheet.Elements<AutoFilter>().Any())
                {
                    var previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.AutoFilter);
                    worksheetPart.Worksheet.InsertAfter(new AutoFilter(), previousElement);
                }

                var autoFilter = worksheetPart.Worksheet.Elements<AutoFilter>().First();
                cm.SetElement(XLWSContentManager.XLWSContents.AutoFilter, autoFilter);

                autoFilter.Reference = xlWorksheet.AutoFilterRange.RangeAddress.ToString();
            }
            else
            {
                worksheetPart.Worksheet.RemoveAllChildren<AutoFilter>();
                cm.SetElement(XLWSContentManager.XLWSContents.AutoFilter, null);
            }

            #endregion

            #region MergeCells

            if ((xlWorksheet).Internals.MergedRanges.Any())
            {
                if (!worksheetPart.Worksheet.Elements<MergeCells>().Any())
                {
                    var previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.MergeCells);
                    worksheetPart.Worksheet.InsertAfter(new MergeCells(), previousElement);
                }

                var mergeCells = worksheetPart.Worksheet.Elements<MergeCells>().First();
                cm.SetElement(XLWSContentManager.XLWSContents.MergeCells, mergeCells);
                mergeCells.RemoveAllChildren<MergeCell>();

                foreach (MergeCell mergeCell in (xlWorksheet).Internals.MergedRanges.Select(
                    m => m.RangeAddress.FirstAddress.ToString() + ":" + m.RangeAddress.LastAddress.ToString()).Select(
                        merged => new MergeCell { Reference = merged }))
                    mergeCells.AppendChild(mergeCell);

                mergeCells.Count = (UInt32)mergeCells.Count();
            }
            else
            {
                worksheetPart.Worksheet.RemoveAllChildren<MergeCells>();
                cm.SetElement(XLWSContentManager.XLWSContents.MergeCells, null);
            }

            #endregion

            #region DataValidations

            if (!xlWorksheet.DataValidations.Any(d => d.IsDirty()))
            {
                worksheetPart.Worksheet.RemoveAllChildren<DataValidations>();
                cm.SetElement(XLWSContentManager.XLWSContents.DataValidations, null);
            }
            else
            {
                if (!worksheetPart.Worksheet.Elements<DataValidations>().Any())
                {
                    var previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.DataValidations);
                    worksheetPart.Worksheet.InsertAfter(new DataValidations(), previousElement);
                }

                var dataValidations = worksheetPart.Worksheet.Elements<DataValidations>().First();
                cm.SetElement(XLWSContentManager.XLWSContents.DataValidations, dataValidations);
                dataValidations.RemoveAllChildren<DataValidation>();
                foreach (IXLDataValidation dv in xlWorksheet.DataValidations)
                {
                    String sequence = dv.Ranges.Aggregate(String.Empty, (current, r) => current + (r.RangeAddress + " "));

                    if (sequence.Length > 0)
                        sequence = sequence.Substring(0, sequence.Length - 1);

                    var dataValidation = new DataValidation
                                             {
                                                 AllowBlank = dv.IgnoreBlanks,
                                                 Formula1 = new Formula1(dv.MinValue),
                                                 Formula2 = new Formula2(dv.MaxValue),
                                                 Type = dv.AllowedValues.ToOpenXml(),
                                                 ShowErrorMessage = dv.ShowErrorMessage,
                                                 Prompt = dv.InputMessage,
                                                 PromptTitle = dv.InputTitle,
                                                 ErrorTitle = dv.ErrorTitle,
                                                 Error = dv.ErrorMessage,
                                                 ShowDropDown = !dv.InCellDropdown,
                                                 ShowInputMessage = dv.ShowInputMessage,
                                                 ErrorStyle = dv.ErrorStyle.ToOpenXml(),
                                                 Operator = dv.Operator.ToOpenXml(),
                                                 SequenceOfReferences =
                                                     new ListValue<StringValue> { InnerText = sequence }
                                             };

                    dataValidations.AppendChild(dataValidation);
                }
                dataValidations.Count = (UInt32)xlWorksheet.DataValidations.Count();
            }

            #endregion

            #region Hyperlinks

            var relToRemove = worksheetPart.HyperlinkRelationships.ToList();
            relToRemove.ForEach(worksheetPart.DeleteReferenceRelationship);
            if (!xlWorksheet.Hyperlinks.Any())
            {
                worksheetPart.Worksheet.RemoveAllChildren<Hyperlinks>();
                cm.SetElement(XLWSContentManager.XLWSContents.Hyperlinks, null);
            }
            else
            {
                if (!worksheetPart.Worksheet.Elements<Hyperlinks>().Any())
                {
                    var previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.Hyperlinks);
                    worksheetPart.Worksheet.InsertAfter(new Hyperlinks(), previousElement);
                }

                var hyperlinks = worksheetPart.Worksheet.Elements<Hyperlinks>().First();
                cm.SetElement(XLWSContentManager.XLWSContents.Hyperlinks, hyperlinks);
                hyperlinks.RemoveAllChildren<Hyperlink>();
                foreach (XLHyperlink hl in xlWorksheet.Hyperlinks)
                {
                    Hyperlink hyperlink;
                    if (hl.IsExternal)
                    {
                        String rId = context.RelIdGenerator.GetNext(RelType.Workbook);
                        hyperlink = new Hyperlink { Reference = hl.Cell.Address.ToString(), Id = rId };
                        worksheetPart.AddHyperlinkRelationship(hl.ExternalAddress, true, rId);
                    }
                    else
                    {
                        hyperlink = new Hyperlink
                                        {
                                            Reference = hl.Cell.Address.ToString(),
                                            Location = hl.InternalAddress,
                                            Display = hl.Cell.GetFormattedString()
                                        };
                    }
                    if (!StringExtensions.IsNullOrWhiteSpace(hl.Tooltip))
                        hyperlink.Tooltip = hl.Tooltip;
                    hyperlinks.AppendChild(hyperlink);
                }
            }

            #endregion

            #region PrintOptions

            if (!worksheetPart.Worksheet.Elements<PrintOptions>().Any())
            {
                var previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.PrintOptions);
                worksheetPart.Worksheet.InsertAfter(new PrintOptions(), previousElement);
            }

            var printOptions = worksheetPart.Worksheet.Elements<PrintOptions>().First();
            cm.SetElement(XLWSContentManager.XLWSContents.PrintOptions, printOptions);

            printOptions.HorizontalCentered = xlWorksheet.PageSetup.CenterHorizontally;
            printOptions.VerticalCentered = xlWorksheet.PageSetup.CenterVertically;
            printOptions.Headings = xlWorksheet.PageSetup.ShowRowAndColumnHeadings;
            printOptions.GridLines = xlWorksheet.PageSetup.ShowGridlines;

            #endregion

            #region PageMargins

            if (!worksheetPart.Worksheet.Elements<PageMargins>().Any())
            {
                var previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.PageMargins);
                worksheetPart.Worksheet.InsertAfter(new PageMargins(), previousElement);
            }

            var pageMargins = worksheetPart.Worksheet.Elements<PageMargins>().First();
            cm.SetElement(XLWSContentManager.XLWSContents.PageMargins, pageMargins);
            pageMargins.Left = xlWorksheet.PageSetup.Margins.Left;
            pageMargins.Right = xlWorksheet.PageSetup.Margins.Right;
            pageMargins.Top = xlWorksheet.PageSetup.Margins.Top;
            pageMargins.Bottom = xlWorksheet.PageSetup.Margins.Bottom;
            pageMargins.Header = xlWorksheet.PageSetup.Margins.Header;
            pageMargins.Footer = xlWorksheet.PageSetup.Margins.Footer;

            #endregion

            #region PageSetup

            if (!worksheetPart.Worksheet.Elements<PageSetup>().Any())
            {
                var previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.PageSetup);
                worksheetPart.Worksheet.InsertAfter(new PageSetup(), previousElement);
            }

            var pageSetup = worksheetPart.Worksheet.Elements<PageSetup>().First();
            cm.SetElement(XLWSContentManager.XLWSContents.PageSetup, pageSetup);

            pageSetup.Orientation = xlWorksheet.PageSetup.PageOrientation.ToOpenXml();
            pageSetup.PaperSize = (UInt32)xlWorksheet.PageSetup.PaperSize;
            pageSetup.BlackAndWhite = xlWorksheet.PageSetup.BlackAndWhite;
            pageSetup.Draft = xlWorksheet.PageSetup.DraftQuality;
            pageSetup.PageOrder = xlWorksheet.PageSetup.PageOrder.ToOpenXml();
            pageSetup.CellComments = xlWorksheet.PageSetup.ShowComments.ToOpenXml();
            pageSetup.Errors = xlWorksheet.PageSetup.PrintErrorValue.ToOpenXml();

            if (xlWorksheet.PageSetup.FirstPageNumber > 0)
            {
                pageSetup.FirstPageNumber = (UInt32)xlWorksheet.PageSetup.FirstPageNumber;
                pageSetup.UseFirstPageNumber = true;
            }
            else
            {
                pageSetup.FirstPageNumber = null;
                pageSetup.UseFirstPageNumber = null;
            }

            if (xlWorksheet.PageSetup.HorizontalDpi > 0)
                pageSetup.HorizontalDpi = (UInt32)xlWorksheet.PageSetup.HorizontalDpi;
            else
                pageSetup.HorizontalDpi = null;

            if (xlWorksheet.PageSetup.VerticalDpi > 0)
                pageSetup.VerticalDpi = (UInt32)xlWorksheet.PageSetup.VerticalDpi;
            else
                pageSetup.VerticalDpi = null;

            if (xlWorksheet.PageSetup.Scale > 0)
            {
                pageSetup.Scale = (UInt32)xlWorksheet.PageSetup.Scale;
                pageSetup.FitToWidth = null;
                pageSetup.FitToHeight = null;
            }
            else
            {
                pageSetup.Scale = null;

                if (xlWorksheet.PageSetup.PagesWide > 0)
                    pageSetup.FitToWidth = (UInt32)xlWorksheet.PageSetup.PagesWide;
                else
                    pageSetup.FitToWidth = 0;

                if (xlWorksheet.PageSetup.PagesTall > 0)
                    pageSetup.FitToHeight = (UInt32)xlWorksheet.PageSetup.PagesTall;
                else
                    pageSetup.FitToHeight = 0;
            }

            #endregion

            #region HeaderFooter

            HeaderFooter headerFooter = worksheetPart.Worksheet.Elements<HeaderFooter>().FirstOrDefault();
            if (headerFooter == null) 
                headerFooter = new HeaderFooter();
            else
                worksheetPart.Worksheet.RemoveAllChildren<HeaderFooter>();

            {
                var previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.HeaderFooter);
                worksheetPart.Worksheet.InsertAfter(headerFooter, previousElement);
                cm.SetElement(XLWSContentManager.XLWSContents.HeaderFooter, headerFooter);
            }
            if (((XLHeaderFooter)xlWorksheet.PageSetup.Header).Changed
                || ((XLHeaderFooter)xlWorksheet.PageSetup.Footer).Changed)
            {
                //var headerFooter = worksheetPart.Worksheet.Elements<HeaderFooter>().First();
                
                headerFooter.RemoveAllChildren();

                headerFooter.ScaleWithDoc = xlWorksheet.PageSetup.ScaleHFWithDocument;
                headerFooter.AlignWithMargins = xlWorksheet.PageSetup.AlignHFWithMargins;
                headerFooter.DifferentFirst = true;
                headerFooter.DifferentOddEven = true;

                var oddHeader = new OddHeader(xlWorksheet.PageSetup.Header.GetText(XLHFOccurrence.OddPages));
                headerFooter.AppendChild(oddHeader);
                var oddFooter = new OddFooter(xlWorksheet.PageSetup.Footer.GetText(XLHFOccurrence.OddPages));
                headerFooter.AppendChild(oddFooter);

                var evenHeader = new EvenHeader(xlWorksheet.PageSetup.Header.GetText(XLHFOccurrence.EvenPages));
                headerFooter.AppendChild(evenHeader);
                var evenFooter = new EvenFooter(xlWorksheet.PageSetup.Footer.GetText(XLHFOccurrence.EvenPages));
                headerFooter.AppendChild(evenFooter);

                var firstHeader = new FirstHeader(xlWorksheet.PageSetup.Header.GetText(XLHFOccurrence.FirstPage));
                headerFooter.AppendChild(firstHeader);
                var firstFooter = new FirstFooter(xlWorksheet.PageSetup.Footer.GetText(XLHFOccurrence.FirstPage));
                headerFooter.AppendChild(firstFooter);

            }

            #endregion

            #region RowBreaks

            if (!worksheetPart.Worksheet.Elements<RowBreaks>().Any())
            {
                var previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.RowBreaks);
                worksheetPart.Worksheet.InsertAfter(new RowBreaks(), previousElement);
            }

            var rowBreaks = worksheetPart.Worksheet.Elements<RowBreaks>().First();

            int rowBreakCount = xlWorksheet.PageSetup.RowBreaks.Count;
            if (rowBreakCount > 0)
            {
                rowBreaks.Count = (UInt32)rowBreakCount;
                rowBreaks.ManualBreakCount = (UInt32)rowBreakCount;
                uint lastRowNum = (UInt32)xlWorksheet.RangeAddress.LastAddress.RowNumber;
                foreach (Break break1 in xlWorksheet.PageSetup.RowBreaks.Select(rb => new Break
                                                                                          {
                                                                                              Id = (UInt32)rb,
                                                                                              Max = lastRowNum,
                                                                                              ManualPageBreak = true
                                                                                          }))
                    rowBreaks.AppendChild(break1);
                cm.SetElement(XLWSContentManager.XLWSContents.RowBreaks, rowBreaks);
            }
            else
            {
                worksheetPart.Worksheet.RemoveAllChildren<RowBreaks>();
                cm.SetElement(XLWSContentManager.XLWSContents.RowBreaks, null);
            }

            #endregion

            #region ColumnBreaks

            if (!worksheetPart.Worksheet.Elements<ColumnBreaks>().Any())
            {
                var previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.ColumnBreaks);
                worksheetPart.Worksheet.InsertAfter(new ColumnBreaks(), previousElement);
            }

            var columnBreaks = worksheetPart.Worksheet.Elements<ColumnBreaks>().First();

            int columnBreakCount = xlWorksheet.PageSetup.ColumnBreaks.Count;
            if (columnBreakCount > 0)
            {
                columnBreaks.Count = (UInt32)columnBreakCount;
                columnBreaks.ManualBreakCount = (UInt32)columnBreakCount;
                uint maxColumnNumber = (UInt32)xlWorksheet.RangeAddress.LastAddress.ColumnNumber;
                foreach (Break break1 in xlWorksheet.PageSetup.ColumnBreaks.Select(cb => new Break
                                                                                             {
                                                                                                 Id = (UInt32)cb,
                                                                                                 Max = maxColumnNumber,
                                                                                                 ManualPageBreak = true
                                                                                             }))
                    columnBreaks.AppendChild(break1);
                cm.SetElement(XLWSContentManager.XLWSContents.ColumnBreaks, columnBreaks);
            }
            else
            {
                worksheetPart.Worksheet.RemoveAllChildren<ColumnBreaks>();
                cm.SetElement(XLWSContentManager.XLWSContents.ColumnBreaks, null);
            }

            #endregion

            #region Drawings

            //worksheetPart.Worksheet.RemoveAllChildren<Drawing>();
            //{
            //    OpenXmlElement previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.Drawing);
            //    worksheetPart.Worksheet.InsertAfter(new Drawing() { Id = String.Format("rId{0}", 1) }, previousElement);
            //}

            //Drawing drawing = worksheetPart.Worksheet.Elements<Drawing>().First();
            //cm.SetElement(XLWSContentManager.XLWSContents.Drawing, drawing);

            #endregion

            #region Tables

            worksheetPart.Worksheet.RemoveAllChildren<TableParts>();
            {
                var previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.TableParts);
                worksheetPart.Worksheet.InsertAfter(new TableParts(), previousElement);
            }

            var tableParts = worksheetPart.Worksheet.Elements<TableParts>().First();
            cm.SetElement(XLWSContentManager.XLWSContents.TableParts, tableParts);

            tableParts.Count = (UInt32)xlWorksheet.Tables.Count();
            foreach (
                TablePart tablePart in
                    from XLTable xlTable in xlWorksheet.Tables select new TablePart { Id = xlTable.RelId })
                tableParts.AppendChild(tablePart);

            #endregion

            #region LegacyDrawing
            //worksheetPart.Worksheet.RemoveAllChildren<LegacyDrawing>();
            //{
            //    if (!StringExtensions.IsNullOrWhiteSpace(xlWorksheet.LegacyDrawingId))
            //    {
            //        var previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.LegacyDrawing);
            //        worksheetPart.Worksheet.InsertAfter(new LegacyDrawing { Id = xlWorksheet.LegacyDrawingId },
            //                                            previousElement);
            //    }
            //}
            #endregion
        }

        private static BooleanValue GetBooleanValue(bool value, bool defaultValue)
        {
            return value == defaultValue ? null : new BooleanValue(value);
        }

        private static void CollapseColumns(Columns columns, Dictionary<uint, Column> sheetColumns)
        {
            UInt32 lastMin = 1;
            Int32 count = sheetColumns.Count;
            var arr = sheetColumns.OrderBy(kp => kp.Key).ToArray();
            // sheetColumns[kp.Key + 1]
            //Int32 i = 0;
            //foreach (KeyValuePair<uint, Column> kp in arr
            //    //.Where(kp => !(kp.Key < count && ColumnsAreEqual(kp.Value, )))
            //    )
            for (int i = 0; i < count; i++)
            {
                var kp = arr[i];
                if (i + 1 != count && ColumnsAreEqual(kp.Value, arr[i + 1].Value)) continue;

                var newColumn = (Column)kp.Value.CloneNode(true);
                newColumn.Min = lastMin;
                uint newColumnMax = newColumn.Max.Value;
                var columnsToRemove =
                    columns.Elements<Column>().Where(co => co.Min >= lastMin && co.Max <= newColumnMax).
                        Select(co => co).ToList();
                columnsToRemove.ForEach(c => columns.RemoveChild(c));

                columns.AppendChild(newColumn);
                lastMin = kp.Key + 1;
                //i++;
            }

        }

        private static double GetColumnWidth(double columnWidth)
        {
            if (columnWidth > 0)
                return columnWidth + ColumnWidthOffset;
            return columnWidth;
        }

        private static void UpdateColumn(Column column, Columns columns, Dictionary<uint, Column> sheetColumnsByMin)
        {
            UInt32 co = column.Min.Value;
            Column newColumn;
            if (!sheetColumnsByMin.ContainsKey(co))
            {
                newColumn = (Column)column.CloneNode(true);
                columns.AppendChild(newColumn);
                sheetColumnsByMin.Add(co, newColumn);
            }
            else
            {
                var existingColumn = sheetColumnsByMin[column.Min.Value];
                newColumn = (Column)existingColumn.CloneNode(true);
                newColumn.Min = column.Min;
                newColumn.Max = column.Max;
                newColumn.Style = column.Style;
                newColumn.Width = column.Width;
                newColumn.CustomWidth = column.CustomWidth;

                if (column.Hidden != null)
                    newColumn.Hidden = true;
                else
                    newColumn.Hidden = null;

                if (column.Collapsed != null)
                    newColumn.Collapsed = true;
                else
                    newColumn.Collapsed = null;

                if (column.OutlineLevel != null && column.OutlineLevel > 0)
                    newColumn.OutlineLevel = (byte)column.OutlineLevel;
                else
                    newColumn.OutlineLevel = null;

                sheetColumnsByMin.Remove(column.Min.Value);
                if (existingColumn.Min + 1 > existingColumn.Max)
                {
                    //existingColumn.Min = existingColumn.Min + 1;
                    //columns.InsertBefore(existingColumn, newColumn);
                    //existingColumn.Remove();
                    columns.RemoveChild(existingColumn);
                    columns.AppendChild(newColumn);
                    sheetColumnsByMin.Add(newColumn.Min.Value, newColumn);
                }
                else
                {
                    //columns.InsertBefore(existingColumn, newColumn);
                    columns.AppendChild(newColumn);
                    sheetColumnsByMin.Add(newColumn.Min.Value, newColumn);
                    existingColumn.Min = existingColumn.Min + 1;
                    sheetColumnsByMin.Add(existingColumn.Min.Value, existingColumn);
                }
            }
        }

        private static bool ColumnsAreEqual(Column left, Column right)
        {
            return
                ((left.Style == null && right.Style == null)
                 || (left.Style != null && right.Style != null && left.Style.Value == right.Style.Value))
                && ((left.Width == null && right.Width == null)
                    || (left.Width != null && right.Width != null && left.Width.Value == right.Width.Value))
                && ((left.Hidden == null && right.Hidden == null)
                    || (left.Hidden != null && right.Hidden != null && left.Hidden.Value == right.Hidden.Value))
                && ((left.Collapsed == null && right.Collapsed == null)
                    ||
                    (left.Collapsed != null && right.Collapsed != null && left.Collapsed.Value == right.Collapsed.Value))
                && ((left.OutlineLevel == null && right.OutlineLevel == null)
                    ||
                    (left.OutlineLevel != null && right.OutlineLevel != null &&
                     left.OutlineLevel.Value == right.OutlineLevel.Value));
        }

        #endregion

        //private void GenerateDrawingsPartContent(DrawingsPart drawingsPart, XLWorksheet worksheet)
        //{
        //    if (drawingsPart.WorksheetDrawing == null)
        //        drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing();

        //    var worksheetDrawing = drawingsPart.WorksheetDrawing;

        //    if (!worksheetDrawing.NamespaceDeclarations.Contains(new KeyValuePair<string, string>("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")))
        //        worksheetDrawing.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
        //    if (!worksheetDrawing.NamespaceDeclarations.Contains(new KeyValuePair<string, string>("a", "http://schemas.openxmlformats.org/drawingml/2006/main")))
        //        worksheetDrawing.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

        //    foreach (var chart in worksheet.Charts.OrderBy(c => c.ZOrder).Select(c => c))
        //    {
        //        Xdr.TwoCellAnchor twoCellAnchor = new Xdr.TwoCellAnchor();
        //        worksheetDrawing.AppendChild(twoCellAnchor);
        //        if (chart.Anchor == XLDrawingAnchor.MoveAndSizeWithCells)
        //            twoCellAnchor.EditAs = Xdr.EditAsValues.TwoCell;
        //        else if (chart.Anchor == XLDrawingAnchor.MoveWithCells)
        //            twoCellAnchor.EditAs = Xdr.EditAsValues.OneCell;
        //        else
        //            twoCellAnchor.EditAs = Xdr.EditAsValues.Absolute;

        //        if (twoCellAnchor.FromMarker == null)
        //            twoCellAnchor.FromMarker = new Xdr.FromMarker();
        //        twoCellAnchor.FromMarker.RowId = new Xdr.RowId((chart.FirstRow - 1).ToString());
        //        twoCellAnchor.FromMarker.RowOffset = new Xdr.RowOffset(chart.FirstRowOffset.ToString());
        //        twoCellAnchor.FromMarker.ColumnId = new Xdr.ColumnId((chart.FirstColumn - 1).ToString());
        //        twoCellAnchor.FromMarker.ColumnOffset = new Xdr.ColumnOffset(chart.FirstColumnOffset.ToString());

        //        if (twoCellAnchor.ToMarker == null)
        //            twoCellAnchor.ToMarker = new Xdr.ToMarker();
        //        twoCellAnchor.ToMarker.RowId = new Xdr.RowId((chart.LastRow - 1).ToString());
        //        twoCellAnchor.ToMarker.RowOffset = new Xdr.RowOffset(chart.LastRowOffset.ToString());
        //        twoCellAnchor.ToMarker.ColumnId = new Xdr.ColumnId((chart.LastColumn - 1).ToString());
        //        twoCellAnchor.ToMarker.ColumnOffset = new Xdr.ColumnOffset(chart.LastColumnOffset.ToString());

        //        Xdr.GraphicFrame graphicFrame = new Xdr.GraphicFrame();
        //        twoCellAnchor.AppendChild(graphicFrame);

        //        if (graphicFrame.NonVisualGraphicFrameProperties == null)
        //            graphicFrame.NonVisualGraphicFrameProperties = new Xdr.NonVisualGraphicFrameProperties();

        //        if (graphicFrame.NonVisualGraphicFrameProperties.NonVisualDrawingProperties == null)
        //            graphicFrame.NonVisualGraphicFrameProperties.NonVisualDrawingProperties = new Xdr.NonVisualDrawingProperties() { Id = (UInt32)chart.Id, Name = chart.Name, Description = chart.Description, Hidden = chart.Hidden };
        //        if (graphicFrame.NonVisualGraphicFrameProperties.NonVisualGraphicFrameDrawingProperties == null)
        //            graphicFrame.NonVisualGraphicFrameProperties.NonVisualGraphicFrameDrawingProperties = new Xdr.NonVisualGraphicFrameDrawingProperties();

        //        if (graphicFrame.Transform == null)
        //            graphicFrame.Transform = new Xdr.Transform();

        //        if (chart.HorizontalFlip)
        //            graphicFrame.Transform.HorizontalFlip = true;
        //        else
        //            graphicFrame.Transform.HorizontalFlip = null;

        //        if (chart.VerticalFlip)
        //            graphicFrame.Transform.VerticalFlip = true;
        //        else
        //            graphicFrame.Transform.VerticalFlip = null;

        //        if (chart.Rotation != 0)
        //            graphicFrame.Transform.Rotation = chart.Rotation;
        //        else
        //            graphicFrame.Transform.Rotation = null;

        //        if (graphicFrame.Transform.Offset == null)
        //            graphicFrame.Transform.Offset = new A.Offset();

        //        graphicFrame.Transform.Offset.X = chart.OffsetX;
        //        graphicFrame.Transform.Offset.Y = chart.OffsetY;

        //        if (graphicFrame.Transform.Extents == null)
        //            graphicFrame.Transform.Extents = new A.Extents();

        //        graphicFrame.Transform.Extents.Cx = chart.ExtentLength;
        //        graphicFrame.Transform.Extents.Cy = chart.ExtentWidth;

        //        if (graphicFrame.Graphic == null)
        //            graphicFrame.Graphic = new A.Graphic();

        //        if (graphicFrame.Graphic.GraphicData == null)
        //            graphicFrame.Graphic.GraphicData = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" };

        //        if (!graphicFrame.Graphic.GraphicData.Elements<C.ChartReference>().Any())
        //        {
        //            C.ChartReference chartReference = new C.ChartReference() { Id = "rId" + chart.Id.ToStringLookup() };
        //            chartReference.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
        //            chartReference.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        //            graphicFrame.Graphic.GraphicData.AppendChild(chartReference);
        //        }

        //        if (!twoCellAnchor.Elements<Xdr.ClientData>().Any())
        //            twoCellAnchor.AppendChild(new Xdr.ClientData());
        //    }
        //}

        //private void GenerateChartPartContent(ChartPart chartPart, XLChart xlChart)
        //{
        //    if (chartPart.ChartSpace == null)
        //        chartPart.ChartSpace = new C.ChartSpace();

        //    C.ChartSpace chartSpace = chartPart.ChartSpace;

        //    if (!chartSpace.NamespaceDeclarations.Contains(new KeyValuePair<string, string>("c", "http://schemas.openxmlformats.org/drawingml/2006/chart")))
        //        chartSpace.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
        //    if (!chartSpace.NamespaceDeclarations.Contains(new KeyValuePair<string, string>("a", "http://schemas.openxmlformats.org/drawingml/2006/main")))
        //        chartSpace.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        //    if (!chartSpace.NamespaceDeclarations.Contains(new KeyValuePair<string, string>("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")))
        //        chartSpace.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        //    if (chartSpace.EditingLanguage == null)
        //        chartSpace.EditingLanguage = new C.EditingLanguage() { Val = CultureInfo.CurrentCulture.Name };
        //    else
        //        chartSpace.EditingLanguage.Val = CultureInfo.CurrentCulture.Name;

        //    C.Chart chart = new C.Chart();
        //    chartSpace.AppendChild(chart);

        //    if (chart.Title == null)
        //        chart.Title = new C.Title();

        //    if (chart.Title.Layout == null)
        //        chart.Title.Layout = new C.Layout();

        //    if (chart.View3D == null)
        //        chart.View3D = new C.View3D();

        //    if (chart.View3D.RightAngleAxes == null)
        //        chart.View3D.RightAngleAxes = new C.RightAngleAxes();

        //    chart.View3D.RightAngleAxes.Val = xlChart.RightAngleAxes;

        //    if (chart.PlotArea == null)
        //        chart.PlotArea = new C.PlotArea();

        //    if (chart.PlotArea.Layout == null)
        //        chart.PlotArea.Layout = new C.Layout();

        //    OpenXmlElement chartElement = GetChartElement(xlChart);

        //    chart.PlotArea.AppendChild(chartElement);

        //    C.CategoryAxis categoryAxis1 = new C.CategoryAxis();
        //    C.AxisId axisId4 = new C.AxisId() { Val = (UInt32Value)71429120U };

        //    C.Scaling scaling1 = new C.Scaling();
        //    C.Orientation orientation1 = new C.Orientation() { Val = C.OrientationValues.MinMax };

        //    scaling1.AppendChild(orientation1);
        //    C.AxisPosition axisPosition1 = new C.AxisPosition() { Val = C.AxisPositionValues.Bottom };
        //    C.TickLabelPosition tickLabelPosition1 = new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo };
        //    C.CrossingAxis crossingAxis1 = new C.CrossingAxis() { Val = (UInt32Value)71432064U };
        //    C.Crosses crosses1 = new C.Crosses() { Val = C.CrossesValues.AutoZero };
        //    C.AutoLabeled autoLabeled1 = new C.AutoLabeled() { Val = true };
        //    C.LabelAlignment labelAlignment1 = new C.LabelAlignment() { Val = C.LabelAlignmentValues.Center };
        //    C.LabelOffset labelOffset1 = new C.LabelOffset() { Val = (UInt16Value)100U };

        //    categoryAxis1.AppendChild(axisId4);
        //    categoryAxis1.AppendChild(scaling1);
        //    categoryAxis1.AppendChild(axisPosition1);
        //    categoryAxis1.AppendChild(tickLabelPosition1);
        //    categoryAxis1.AppendChild(crossingAxis1);
        //    categoryAxis1.AppendChild(crosses1);
        //    categoryAxis1.AppendChild(autoLabeled1);
        //    categoryAxis1.AppendChild(labelAlignment1);
        //    categoryAxis1.AppendChild(labelOffset1);

        //    C.ValueAxis valueAxis1 = new C.ValueAxis();
        //    C.AxisId axisId5 = new C.AxisId() { Val = (UInt32Value)71432064U };

        //    C.Scaling scaling2 = new C.Scaling();
        //    C.Orientation orientation2 = new C.Orientation() { Val = C.OrientationValues.MinMax };

        //    scaling2.AppendChild(orientation2);
        //    C.AxisPosition axisPosition2 = new C.AxisPosition() { Val = C.AxisPositionValues.Left };
        //    C.MajorGridlines majorGridlines1 = new C.MajorGridlines();
        //    C.NumberingFormat numberingFormat1 = new C.NumberingFormat() { FormatCode = "General", SourceLinked = true };
        //    C.TickLabelPosition tickLabelPosition2 = new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo };
        //    C.CrossingAxis crossingAxis2 = new C.CrossingAxis() { Val = (UInt32Value)71429120U };
        //    C.Crosses crosses2 = new C.Crosses() { Val = C.CrossesValues.AutoZero };
        //    C.CrossBetween crossBetween1 = new C.CrossBetween() { Val = C.CrossBetweenValues.Between };

        //    valueAxis1.AppendChild(axisId5);
        //    valueAxis1.AppendChild(scaling2);
        //    valueAxis1.AppendChild(axisPosition2);
        //    valueAxis1.AppendChild(majorGridlines1);
        //    valueAxis1.AppendChild(numberingFormat1);
        //    valueAxis1.AppendChild(tickLabelPosition2);
        //    valueAxis1.AppendChild(crossingAxis2);
        //    valueAxis1.AppendChild(crosses2);
        //    valueAxis1.AppendChild(crossBetween1);

        //    plotArea.AppendChild(bar3DChart1);
        //    plotArea.AppendChild(categoryAxis1);
        //    plotArea.AppendChild(valueAxis1);

        //    C.Legend legend1 = new C.Legend();
        //    C.LegendPosition legendPosition1 = new C.LegendPosition() { Val = C.LegendPositionValues.Right };
        //    C.Layout layout3 = new C.Layout();

        //    legend1.AppendChild(legendPosition1);
        //    legend1.AppendChild(layout3);
        //    C.PlotVisibleOnly plotVisibleOnly1 = new C.PlotVisibleOnly() { Val = true };

        //    chart.AppendChild(legend1);
        //    chart.AppendChild(plotVisibleOnly1);

        //    C.PrintSettings printSettings1 = new C.PrintSettings();
        //    C.HeaderFooter headerFooter1 = new C.HeaderFooter();
        //    C.PageMargins pageMargins4 = new C.PageMargins() { Left = 0.70000000000000018D, Right = 0.70000000000000018D, Top = 0.75000000000000022D, Bottom = 0.75000000000000022D, Header = 0.3000000000000001D, Footer = 0.3000000000000001D };
        //    C.PageSetup pageSetup1 = new C.PageSetup();

        //    printSettings1.AppendChild(headerFooter1);
        //    printSettings1.AppendChild(pageMargins4);
        //    printSettings1.AppendChild(pageSetup1);

        //    chartSpace.AppendChild(printSettings1);

        //}

        //private OpenXmlElement GetChartElement(XLChart xlChart)
        //{
        //    if (xlChart.ChartTypeCategory == XLChartTypeCategory.Bar3D)
        //        return GetBar3DChart(xlChart);
        //    else
        //        return null;
        //}

        //private OpenXmlElement GetBar3DChart(XLChart xlChart)
        //{

        //    C.Bar3DChart bar3DChart = new C.Bar3DChart();
        //    bar3DChart.BarDirection = new C.BarDirection() { Val = GetBarDirection(xlChart) };
        //    bar3DChart.BarGrouping = new C.BarGrouping() { Val = GetBarGrouping(xlChart) };

        //    C.BarChartSeries barChartSeries = new C.BarChartSeries();
        //    barChartSeries.Index = new C.Index() { Val = (UInt32Value)0U };
        //    barChartSeries.Order = new C.Order() { Val = (UInt32Value)0U };

        //    C.SeriesText seriesText1 = new C.SeriesText();

        //    C.StringReference stringReference1 = new C.StringReference();
        //    C.Formula formula1 = new C.Formula();
        //    formula1.Text = "Sheet1!$B$1";

        //    stringReference1.AppendChild(formula1);

        //    seriesText1.AppendChild(stringReference1);

        //    C.CategoryAxisData categoryAxisData1 = new C.CategoryAxisData();

        //    C.StringReference stringReference2 = new C.StringReference();
        //    C.Formula formula2 = new C.Formula();
        //    formula2.Text = "Sheet1!$A$2:$A$3";

        //    C.StringCache stringCache2 = new C.StringCache();
        //    C.PointCount pointCount2 = new C.PointCount() { Val = (UInt32Value)2U };

        //    C.StringPoint stringPoint2 = new C.StringPoint() { Index = (UInt32Value)0U };
        //    C.NumericValue numericValue2 = new C.NumericValue();
        //    numericValue2.Text = "A";

        //    stringPoint2.AppendChild(numericValue2);

        //    C.StringPoint stringPoint3 = new C.StringPoint() { Index = (UInt32Value)1U };
        //    C.NumericValue numericValue3 = new C.NumericValue();
        //    numericValue3.Text = "B";

        //    stringPoint3.AppendChild(numericValue3);

        //    stringCache2.AppendChild(pointCount2);
        //    stringCache2.AppendChild(stringPoint2);
        //    stringCache2.AppendChild(stringPoint3);

        //    stringReference2.AppendChild(formula2);
        //    stringReference2.AppendChild(stringCache2);

        //    categoryAxisData1.AppendChild(stringReference2);

        //    C.Values values1 = new C.Values();

        //    C.NumberReference numberReference1 = new C.NumberReference();
        //    C.Formula formula3 = new C.Formula();
        //    formula3.Text = "Sheet1!$B$2:$B$3";

        //    C.NumberingCache numberingCache1 = new C.NumberingCache();
        //    C.FormatCode formatCode1 = new C.FormatCode();
        //    formatCode1.Text = "General";
        //    C.PointCount pointCount3 = new C.PointCount() { Val = (UInt32Value)2U };

        //    C.NumericPoint numericPoint1 = new C.NumericPoint() { Index = (UInt32Value)0U };
        //    C.NumericValue numericValue4 = new C.NumericValue();
        //    numericValue4.Text = "5";

        //    numericPoint1.AppendChild(numericValue4);

        //    C.NumericPoint numericPoint2 = new C.NumericPoint() { Index = (UInt32Value)1U };
        //    C.NumericValue numericValue5 = new C.NumericValue();
        //    numericValue5.Text = "10";

        //    numericPoint2.AppendChild(numericValue5);

        //    numberingCache1.AppendChild(formatCode1);
        //    numberingCache1.AppendChild(pointCount3);
        //    numberingCache1.AppendChild(numericPoint1);
        //    numberingCache1.AppendChild(numericPoint2);

        //    numberReference1.AppendChild(formula3);
        //    numberReference1.AppendChild(numberingCache1);

        //    values1.AppendChild(numberReference1);

        //    barChartSeries.AppendChild(index1);
        //    barChartSeries.AppendChild(order1);
        //    barChartSeries.AppendChild(seriesText1);
        //    barChartSeries.AppendChild(categoryAxisData1);
        //    barChartSeries.AppendChild(values1);
        //    C.Shape shape1 = new C.Shape() { Val = C.ShapeValues.Box };
        //    C.AxisId axisId1 = new C.AxisId() { Val = (UInt32Value)71429120U };
        //    C.AxisId axisId2 = new C.AxisId() { Val = (UInt32Value)71432064U };
        //    C.AxisId axisId3 = new C.AxisId() { Val = (UInt32Value)0U };

        //    bar3DChart.AppendChild(barChartSeries);
        //    bar3DChart.AppendChild(shape1);
        //    bar3DChart.AppendChild(axisId1);
        //    bar3DChart.AppendChild(axisId2);
        //    bar3DChart.AppendChild(axisId3);

        //    return bar3DChart;
        //}

        //private C.BarGroupingValues GetBarGrouping(XLChart xlChart)
        //{
        //    if (xlChart.BarGrouping == XLBarGrouping.Clustered)
        //        return C.BarGroupingValues.Clustered;
        //    else if (xlChart.BarGrouping == XLBarGrouping.Percent)
        //        return C.BarGroupingValues.PercentStacked;
        //    else if (xlChart.BarGrouping == XLBarGrouping.Stacked)
        //        return C.BarGroupingValues.Stacked;
        //    else
        //        return C.BarGroupingValues.Standard;
        //}

        //private C.BarDirectionValues GetBarDirection(XLChart xlChart)
        //{
        //    if (xlChart.BarOrientation == XLBarOrientation.Vertical)
        //        return C.BarDirectionValues.Column;
        //    else
        //        return C.BarDirectionValues.Bar;
        //}
        //--

        private static void GeneratePivotTables(WorkbookPart workbookPart, WorksheetPart worksheetPart, XLWorksheet xlWorksheet,
                                                         SaveContext context)
        {
            foreach (var pt in xlWorksheet.PivotTables)
            {
                string ptCdp = context.RelIdGenerator.GetNext(RelType.Workbook);

                var pivotTableCacheDefinitionPart = workbookPart.AddNewPart<PivotTableCacheDefinitionPart>(ptCdp);
                GeneratePivotTableCacheDefinitionPartContent(pivotTableCacheDefinitionPart, pt);

                var pivotCaches = new PivotCaches();
                var pivotCache = new PivotCache { CacheId = 0U, Id = ptCdp };

                pivotCaches.AppendChild(pivotCache);

                workbookPart.Workbook.AppendChild(pivotCaches);

                var pivotTablePart = worksheetPart.AddNewPart<PivotTablePart>(context.RelIdGenerator.GetNext(RelType.Workbook));
                GeneratePivotTablePartContent(pivotTablePart, pt);

                pivotTablePart.AddPart(pivotTableCacheDefinitionPart, context.RelIdGenerator.GetNext(RelType.Workbook));
            }
        }

        // Generates content of pivotTableCacheDefinitionPart
        private static void GeneratePivotTableCacheDefinitionPartContent(PivotTableCacheDefinitionPart pivotTableCacheDefinitionPart, IXLPivotTable pt)
        {
            IXLRange source = pt.SourceRange;

            var pivotCacheDefinition = new PivotCacheDefinition
            {
                Id = "rId1",
                SaveData = pt.SaveSourceData,
                RefreshOnLoad = true //pt.RefreshDataOnOpen
            };
            if (pt.ItemsToRetainPerField == XLItemsToRetain.None)
                pivotCacheDefinition.MissingItemsLimit = 0U;
            else if (pt.ItemsToRetainPerField == XLItemsToRetain.Max)
                pivotCacheDefinition.MissingItemsLimit = ExcelHelper.MaxRowNumber;

            pivotCacheDefinition.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            var cacheSource = new CacheSource { Type = SourceValues.Worksheet };
            cacheSource.AppendChild(new WorksheetSource { Name = source.ToString() });

            var cacheFields = new CacheFields();

            foreach (var c in source.Columns())
            {
                var columnNumber = c.ColumnNumber();
                var columnName = c.FirstCell().Value.ToString();
                var xlpf = pt.Fields.Add(columnName);

                var field = pt.RowLabels.Union(pt.ColumnLabels).Union(pt.ReportFilters).Where(f => f.SourceName == columnName).FirstOrDefault();
                if (field != null)
                {
                    xlpf.CustomName = field.CustomName;
                    xlpf.Subtotals.AddRange(field.Subtotals);
                }

                var sharedItems = new SharedItems();

                var onlyNumbers = !source.Cells().Any(cell => cell.Address.ColumnNumber == columnNumber && cell.Address.RowNumber > source.FirstRow().RowNumber() && cell.DataType != XLCellValues.Number);
                if (onlyNumbers)
                {
                    sharedItems = new SharedItems { ContainsSemiMixedTypes = false, ContainsString = false, ContainsNumber = true };
                }
                else
                {
                    foreach (var cellValue in source.Cells().Where(cell =>
                                                                   cell.Address.ColumnNumber == columnNumber &&
                                                                   cell.Address.RowNumber > source.FirstRow().RowNumber()).Select(cell => cell.Value.ToString())
                                                                   .Where(cellValue => !xlpf.SharedStrings.Contains(cellValue)))
                    {
                        xlpf.SharedStrings.Add(cellValue);
                    }

                    foreach (var li in xlpf.SharedStrings)
                    {
                        sharedItems.AppendChild(new StringItem { Val = li });
                    }
                }

                var cacheField = new CacheField { Name = xlpf.SourceName };
                cacheField.AppendChild(sharedItems);
                cacheFields.AppendChild(cacheField);
            }

            pivotCacheDefinition.AppendChild(cacheSource);
            pivotCacheDefinition.AppendChild(cacheFields);

            pivotTableCacheDefinitionPart.PivotCacheDefinition = pivotCacheDefinition;

            var pivotTableCacheRecordsPart = pivotTableCacheDefinitionPart.AddNewPart<PivotTableCacheRecordsPart>("rId1");

            var pivotCacheRecords = new PivotCacheRecords();
            pivotCacheRecords.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            pivotTableCacheRecordsPart.PivotCacheRecords = pivotCacheRecords;

        }

        // Generates content of pivotTablePart
        private static void GeneratePivotTablePartContent(PivotTablePart pivotTablePart1, IXLPivotTable pt)
        {
            var pivotTableDefinition = new PivotTableDefinition
                {
                    Name = pt.Name,
                    CacheId = 0U,
                    DataCaption = "Values",
                    MergeItem = GetBooleanValue(pt.MergeAndCenterWithLabels, true),
                    Indent = Convert.ToUInt32(pt.RowLabelIndent),
                    PageOverThenDown = (pt.FilterAreaOrder == XLFilterAreaOrder.OverThenDown),
                    PageWrap = Convert.ToUInt32(pt.FilterFieldsPageWrap),
                    ShowError = String.IsNullOrEmpty(pt.ErrorValueReplacement),
                    UseAutoFormatting = GetBooleanValue(pt.AutofitColumns, true),
                    PreserveFormatting = GetBooleanValue(pt.PreserveCellFormatting, true),
                    RowGrandTotals = GetBooleanValue(pt.ShowGrandTotalsRows, true),
                    ColumnGrandTotals = GetBooleanValue(pt.ShowGrandTotalsColumns, true),
                    SubtotalHiddenItems = GetBooleanValue(pt.FilteredItemsInSubtotals, true),
                    MultipleFieldFilters = GetBooleanValue(pt.AllowMultipleFilters, true),
                    CustomListSort = GetBooleanValue(pt.UseCustomListsForSorting, true),
                    ShowDrill = GetBooleanValue(pt.ShowExpandCollapseButtons, true),
                    ShowDataTips = GetBooleanValue(pt.ShowContextualTooltips, true),
                    ShowMemberPropertyTips = GetBooleanValue(pt.ShowPropertiesInTooltips, true),
                    ShowHeaders = GetBooleanValue(pt.DisplayCaptionsAndDropdowns, true),
                    GridDropZones = GetBooleanValue(pt.ClassicPivotTableLayout, true),
                    ShowEmptyRow = GetBooleanValue(pt.ShowEmptyItemsOnRows, true),
                    ShowEmptyColumn = GetBooleanValue(pt.ShowEmptyItemsOnColumns, true),
                    ShowItems = GetBooleanValue(pt.DisplayItemLabels, true),
                    FieldListSortAscending = GetBooleanValue(pt.SortFieldsAtoZ, true),
                    PrintDrill = GetBooleanValue(pt.PrintExpandCollapsedButtons, true),
                    ItemPrintTitles = GetBooleanValue(pt.RepeatRowLabels, true),
                    FieldPrintTitles = GetBooleanValue(pt.PrintTitles, true),
                    EnableDrill = GetBooleanValue(pt.EnableShowDetails, true)
                };

            if (pt.EmptyCellReplacement != null)
            {
                pivotTableDefinition.ShowMissing = true;
                pivotTableDefinition.MissingCaption = pt.EmptyCellReplacement;
            }
            else
            {
                pivotTableDefinition.ShowMissing = false;
            }

            if (pt.ErrorValueReplacement != null)
            {
                pivotTableDefinition.ShowError = true;
                pivotTableDefinition.ErrorCaption = pt.ErrorValueReplacement;
            }
            else
            {
                pivotTableDefinition.ShowError = false;
            }

            var location = new Location { Reference = pt.TargetCell.Address.ToString(), FirstHeaderRow = 1U, FirstDataRow = 1U, FirstDataColumn = 1U };


            var rowFields = new RowFields();
            var columnFields = new ColumnFields();
            var rowItems = new RowItems();
            var columnItems = new ColumnItems();
            var pageFields = new PageFields { Count = (uint)pt.ReportFilters.Count()};

            var pivotFields = new PivotFields { Count = Convert.ToUInt32(pt.SourceRange.ColumnCount()) };
            foreach (var xlpf in pt.Fields)
            {
                var pf = new PivotField { ShowAll = false, Name = xlpf.CustomName };

                

                if (pt.RowLabels.Where(p => p.SourceName == xlpf.SourceName).FirstOrDefault() != null)
                {
                    pf.Axis = PivotTableAxisValues.AxisRow;

                    var f = new DocumentFormat.OpenXml.Spreadsheet.Field { Index = pt.Fields.IndexOf(xlpf) };
                    rowFields.AppendChild(f);

                    for (int i = 0; i < xlpf.SharedStrings.Count; i++)
                    {
                        var rowItem = new RowItem();
                        rowItem.AppendChild(new MemberPropertyIndex { Val = i });
                       rowItems.AppendChild(rowItem);
                    }

                    var rowItemTotal = new RowItem { ItemType = ItemValues.Grand };
                    rowItemTotal.AppendChild(new MemberPropertyIndex());
                    rowItems.AppendChild(rowItemTotal);


                }
                else if (pt.ColumnLabels.Where(p => p.SourceName == xlpf.SourceName).FirstOrDefault() != null)
                {
                    pf.Axis = PivotTableAxisValues.AxisColumn;

                    var f = new DocumentFormat.OpenXml.Spreadsheet.Field { Index = pt.Fields.IndexOf(xlpf) };
                    columnFields.AppendChild(f);

                    for (int i = 0; i < xlpf.SharedStrings.Count; i++)
                    {
                        var rowItem = new RowItem();
                        rowItem.AppendChild(new MemberPropertyIndex { Val = i });
                        columnItems.AppendChild(rowItem);
                }

                    var rowItemTotal = new RowItem { ItemType = ItemValues.Grand };
                    rowItemTotal.AppendChild(new MemberPropertyIndex());
                    columnItems.AppendChild(rowItemTotal);


                }
                else if (pt.ReportFilters.Where(p => p.SourceName == xlpf.SourceName).FirstOrDefault() != null)
                {
                    location.ColumnsPerPage = 1;
                    location.RowPageCount = 1;
                    pf.Axis = PivotTableAxisValues.AxisPage;
                    pageFields.AppendChild(new PageField {Hierarchy = -1, Field = pt.Fields.IndexOf(xlpf)});
                } 
                else if (pt.Values.Where(p => p.CustomName == xlpf.SourceName).FirstOrDefault() != null)
                {
                    pf.DataField = true;
                }
                
                var fieldItems = new Items();

                if (xlpf.SharedStrings.Count > 0)
                {
                    for (uint i = 0; i < xlpf.SharedStrings.Count; i++)
                    {
                        fieldItems.AppendChild(new Item { Index = i });
                    }  
                }

                if (xlpf.Subtotals.Count > 0)
                {
                    foreach (var subtotal in xlpf.Subtotals)
                    {
                        var itemSubtotal = new Item();
                        switch (subtotal)
                        {
                            case XLSubtotalFunction.Average:
                                pf.AverageSubTotal = true;
                                itemSubtotal.ItemType = ItemValues.Average;
                                break;
                            case XLSubtotalFunction.Count:
                                pf.CountASubtotal = true;
                                itemSubtotal.ItemType = ItemValues.CountA;
                                break;
                            case XLSubtotalFunction.CountNumbers:
                                pf.CountSubtotal = true;
                                itemSubtotal.ItemType = ItemValues.Count;
                                break;
                            case XLSubtotalFunction.Maximum:
                                pf.MaxSubtotal = true;
                                itemSubtotal.ItemType = ItemValues.Maximum;
                                break;
                            case XLSubtotalFunction.Minimum:
                                pf.MinSubtotal = true;
                                itemSubtotal.ItemType = ItemValues.Minimum;
                                break;
                            case XLSubtotalFunction.PopulationStandardDeviation:
                                pf.ApplyStandardDeviationPInSubtotal = true;
                                itemSubtotal.ItemType = ItemValues.StandardDeviationP;
                                break;
                            case XLSubtotalFunction.PopulationVariance:
                                pf.ApplyVariancePInSubtotal = true;
                                itemSubtotal.ItemType = ItemValues.VarianceP;
                                break;
                            case XLSubtotalFunction.Product:
                                pf.ApplyProductInSubtotal = true;
                                itemSubtotal.ItemType = ItemValues.Product;
                                break;
                            case XLSubtotalFunction.StandardDeviation:
                                pf.ApplyStandardDeviationInSubtotal = true;
                                itemSubtotal.ItemType = ItemValues.StandardDeviation;
                                break;
                            case XLSubtotalFunction.Sum:
                                pf.SumSubtotal = true;
                                itemSubtotal.ItemType = ItemValues.Sum;
                                break;
                            case XLSubtotalFunction.Variance:
                                pf.ApplyVarianceInSubtotal = true;
                                itemSubtotal.ItemType = ItemValues.Variance;
                                break;
                        }
                        fieldItems.AppendChild(itemSubtotal);
                    }
                }
                else
                {
                    fieldItems.AppendChild(new Item { ItemType = ItemValues.Default });
                }

                pf.AppendChild(fieldItems);
                pivotFields.AppendChild(pf);
            }

            pivotTableDefinition.AppendChild(location);
            pivotTableDefinition.AppendChild(pivotFields);

            if (pt.RowLabels.Count() > 0)
            {
                pivotTableDefinition.AppendChild(rowFields);
            }
            else
            {
                rowItems.AppendChild(new RowItem());
            }
            pivotTableDefinition.AppendChild(rowItems);

            if (pt.ColumnLabels.Count() == 0)
            {
                columnItems.AppendChild(new RowItem());
                pivotTableDefinition.AppendChild(columnItems);
            }
            else
            {
                pivotTableDefinition.AppendChild(columnFields);
                pivotTableDefinition.AppendChild(columnItems);
            }

            if (pt.ReportFilters.Count() > 0)
            {
                pivotTableDefinition.AppendChild(pageFields);
            }


            var dataFields = new DataFields();
            foreach (var value in pt.Values)
            {
                var sourceColumn = pt.SourceRange.Columns().Where(c => c.Cell(1).Value.ToString() == value.SourceName).FirstOrDefault();
                if (sourceColumn == null) continue;

                var df = new DataField
                             {
                                 Name = value.SourceName,
                                 Field = (UInt32)sourceColumn.ColumnNumber() - 1,
                                 Subtotal = value.SummaryFormula.ToOpenXml(),
                                 ShowDataAs = value.Calculation.ToOpenXml(),
                                 NumberFormatId = (UInt32)value.NumberFormat.NumberFormatId
                             };

                if (!String.IsNullOrEmpty(value.BaseField))
                {
                    var baseField = pt.SourceRange.Columns().Where(c => c.Cell(1).Value.ToString() == value.BaseField).FirstOrDefault();
                    if (baseField != null)
                        df.BaseField = baseField.ColumnNumber() - 1;
                }
                else
                {
                    df.BaseField = 0;
                }

                if (value.CalculationItem == XLPivotCalculationItem.Previous)
                    df.BaseItem = 1048828U;
                else if (value.CalculationItem == XLPivotCalculationItem.Next)
                    df.BaseItem = 1048829U;
                else
                    df.BaseItem = 0U;


                dataFields.AppendChild(df);
            }
            pivotTableDefinition.AppendChild(dataFields);

            pivotTableDefinition.AppendChild(new PivotTableStyle { Name = Enum.GetName(typeof(XLPivotTableTheme), pt.Theme), ShowRowHeaders = pt.ShowRowHeaders, ShowColumnHeaders = pt.ShowColumnHeaders, ShowRowStripes = pt.ShowRowStripes, ShowColumnStripes = pt.ShowColumnStripes });

            #region Excel 2010 Features
            
            var pivotTableDefinitionExtensionList = new PivotTableDefinitionExtensionList();

            var pivotTableDefinitionExtension = new PivotTableDefinitionExtension { Uri = "{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}" };
            pivotTableDefinitionExtension.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");

            var pivotTableDefinition2 = new DocumentFormat.OpenXml.Office2010.Excel.PivotTableDefinition { EnableEdit = pt.EnableCellEditing, HideValuesRow = !pt.ShowValuesRow };
            pivotTableDefinition2.AddNamespaceDeclaration("xm", "http://schemas.microsoft.com/office/excel/2006/main");

            pivotTableDefinitionExtension.AppendChild(pivotTableDefinition2);

            pivotTableDefinitionExtensionList.AppendChild(pivotTableDefinitionExtension);
            pivotTableDefinition.AppendChild(pivotTableDefinitionExtensionList);
            
            #endregion

            pivotTablePart1.PivotTableDefinition = pivotTableDefinition;
        }


        private static void GenerateWorksheetCommentsPartContent(WorksheetCommentsPart worksheetCommentsPart, XLWorksheet xlWorksheet)
        {
            Comments comments = new Comments();
            CommentList commentList = new CommentList();
            var authorsDict = new Dictionary<String, Int32>();
            foreach (var c in xlWorksheet.Internals.CellsCollection.GetCells(c=>c.HasComment))
            {
                Comment comment = new Comment() { Reference = c.Address.ToStringRelative() };
                String authorName = StringExtensions.IsNullOrWhiteSpace(c.Comment.Author)
                                        ? Environment.UserName
                                        : c.Comment.Author;

                    Int32 authorId;
                    if (!authorsDict.TryGetValue(authorName, out authorId))
                    {
                        authorId = authorsDict.Count;
                        authorsDict.Add(authorName, authorId);
                    }
                    comment.AuthorId = (UInt32)authorId;

                CommentText commentText = new CommentText();
                foreach (var rt in c.Comment)
                {
                    commentText.Append(GetRun(rt));
                }

                comment.Append(commentText);
                commentList.Append(comment);
            }

            Authors authors = new Authors();
            foreach (Author author in authorsDict.Select(a => new Author() {Text = a.Key}))
            {
                authors.Append(author);
            }
            comments.Append(authors);
            comments.Append(commentList);

            worksheetCommentsPart.Comments = comments;
        }

        // Generates content of vmlDrawingPart1.
        private static void GenerateVmlDrawingPartContent(VmlDrawingPart vmlDrawingPart, XLWorksheet xlWorksheet, SaveContext context)
        {

            #region Office VML
            // <xml xmlns:v='urn:schemas-microsoft-com:vml' 
            //     xmlns:o='urn:schemas-microsoft-com:office:office' 
            //     xmlns:x='urn:schemas-microsoft-com:office:excel'>

            //     <o:shapelayout v:ext='edit'>
            //     <o:idmap v:ext='edit' data='1'/>
            //     </o:shapelayout>

            //     <!-- SINGLE SHAPE TYPE -->
            //     <v:shapetype id='_x0000_t202' coordsize='21600,21600' o:spt='202'  path='m,l,21600r21600,l21600,xe'>
            //     <v:stroke joinstyle='miter'/>
            //     <v:path gradientshapeok='t' o:connecttype='rect'/>
            //     </v:shapetype>
            //     <!-- /// end -->

            //     <!-- ONE SHAPE PER EACH CELL REFERS to SINGLE SHAPE TYPE above -->
            //     <v:shape id='_x0000_s1026' type='#{0}' style='visibility:hidden' fillcolor='#ffffe1' o:insetmode='auto'>
            //         <v:fill color2='#ffffe1'/>
            //         <v:shadow on='t' color='black' obscured='t'/>
            //         <v:path o:connecttype='none'/>
            //         <v:textbox style='mso-direction-alt:auto'>
            //             <div style='text-align:left'></div>
            //         </v:textbox>
            //         <x:ClientData ObjectType='Note'>
            //             <x:Anchor> {leftCol}, 15, {topRow}, 4, {rightCol}, 10, {bottomRow}, 1</x:Anchor>
            //             <x:Row>{rowIndex}</x:Row>
            //             <x:Column>{colIndex}</x:Column>
            //         </x:ClientData>
            //     </v:shape>  
            //</xml>
            #endregion

            System.Xml.XmlTextWriter writer = new System.Xml.XmlTextWriter(vmlDrawingPart.GetStream(System.IO.FileMode.Create), System.Text.Encoding.UTF8);
            writer.WriteStartElement("xml");

            // o:shapelayout
            new Vml.Office.ShapeLayout(
                new Vml.Office.ShapeIdMap()
                {
                    Extension = Vml.ExtensionHandlingBehaviorValues.Edit,
                    Data = "1"
                }
                ) { Extension = Vml.ExtensionHandlingBehaviorValues.Edit }
                    .WriteTo(writer);

            const string shapeTypeId = "_x0000_t202"; // arbitrary, assigned by office

            // v:shapetype
            new Vml.Shapetype(
                new Vml.Stroke() { JoinStyle = Vml.StrokeJoinStyleValues.Miter },
                new Vml.Path() { AllowGradientShape = true, ConnectionPointType = Vml.Office.ConnectValues.Rectangle }
                )
            {
                Id = shapeTypeId,
                CoordinateSize = "21600,21600",
                OptionalNumber = 202,
                EdgePath = "m,l,21600r21600,l21600,xe",
            }
                    .WriteTo(writer);

            // v:shape
            var cellWithComments = xlWorksheet.Internals.CellsCollection.GetCells().Where(c => c.HasComment);

            foreach (XLCell c in cellWithComments)
            {
                GenerateShape(c, shapeTypeId).WriteTo(writer);
            }

            writer.Flush();
            writer.Close();

        }

        // VML Shape for Comment
        private static Vml.Shape GenerateShape(XLCell c, string shapeTypeId)
        {

            #region Office VML
            //<v:shape id='_x0000_s1026' type='#{0}' style='visibility:hidden' fillcolor='#ffffe1' o:insetmode='auto'>
            //    <v:fill color2='#ffffe1'/>
            //    <v:shadow on='t' color='black' obscured='t'/>
            //    <v:path o:connecttype='none'/>
            //    <v:textbox style='mso-direction-alt:auto'>
            //        <div style='text-align:left'></div>
            //    </v:textbox>
            //    <x:ClientData ObjectType='Note'>
            //        <x:Anchor> {leftCol}, 15, {topRow}, 4, {rightCol}, 10, {bottomRow}, 1</x:Anchor>
            //        <x:Row>{rowIndex}</x:Row>
            //        <x:Column>{colIndex}</x:Column>
            //    </x:ClientData>
            //</v:shape>
            #endregion

            // Limitaion: Most of the shape properties hard coded.

           
            var rowNumber = c.Address.RowNumber;
            var columnNumber = c.Address.ColumnNumber;
            
            var leftCol = columnNumber; // always right next to column
            var leftOffset = 15;
            var topRow = rowNumber == 1 ? rowNumber - 1 : rowNumber - 2;    // -1 : zero based index, -2 : moved up
            var topOffset = rowNumber == 1 ? 2 : 9;    // on first row, comment is 2px down, on any other is 15 px up
            var rightCol = leftCol + c.Comment.Style.Size.Width;            
            var rightOffset = 15;
            var bottomRow = topRow + c.Comment.Style.Size.Height;
            var bottomOffset = rowNumber == 1 ? 2 : 9;

            var shapeId = string.Format("_x0000_s{0}", c.GetHashCode().ToString()); // Unique per cell, e.g.: "_x0000_s1026"

            return new Vml.Shape(
                new Vml.Fill { Color2 = "#" + c.Comment.Style.ColorsAndLines.FillColor.Color.ToHex().Substring(2) },
                new Vml.Shadow() { On = true, Color = "black", Obscured = true },
                new Vml.Path() { ConnectionPointType = Vml.Office.ConnectValues.None },
                new Vml.TextBox( /* <div style='text-align:left'></div> */ ) { Style = "mso-direction-alt:auto" },
                new Vml.Spreadsheet.ClientData(
                    new Vml.Spreadsheet.MoveWithCells("False"),  // counterintuitive
                    new Vml.Spreadsheet.ResizeWithCells("False"), // counterintuitive
                    new Vml.Spreadsheet.Anchor() { Text = string.Format(" {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}", leftCol, leftOffset, topRow, topOffset, rightCol, rightOffset, bottomRow, bottomOffset) },
                    new Vml.Spreadsheet.AutoFill("False"),
                    new Vml.Spreadsheet.CommentRowTarget() { Text = (rowNumber - 1).ToString() },
                    new Vml.Spreadsheet.CommentColumnTarget() { Text = (columnNumber - 1).ToString() }
                    ) { ObjectType = Vml.Spreadsheet.ObjectValues.Note }
                )
                {
                    Id = shapeId,
                    Type = "#" + shapeTypeId,
                    Style = "visibility:hidden",
                    FillColor = "#" + c.Comment.Style.ColorsAndLines.FillColor.Color.ToHex().Substring(2),
                    InsetMode = Vml.Office.InsetMarginValues.Auto
                };
        }
    }
}