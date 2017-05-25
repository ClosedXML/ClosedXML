﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Vml.Spreadsheet;
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
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using System.Xml;
using System.Xml.Linq;
using System.Text;
using ClosedXML.Utils;
using Anchor = DocumentFormat.OpenXml.Vml.Spreadsheet.Anchor;
using Field = DocumentFormat.OpenXml.Spreadsheet.Field;
using Run = DocumentFormat.OpenXml.Spreadsheet.Run;
using RunProperties = DocumentFormat.OpenXml.Spreadsheet.RunProperties;
using VerticalTextAlignment = DocumentFormat.OpenXml.Spreadsheet.VerticalTextAlignment;
using System.Threading;

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

        private bool Validate(SpreadsheetDocument package)
        {
            var backupCulture = Thread.CurrentThread.CurrentCulture;

            IEnumerable<ValidationErrorInfo> errors;
            try
            {
                Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
                var validator = new OpenXmlValidator();
                errors = validator.Validate(package);
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = backupCulture;
            }

            if (errors.Any())
            {
                var message = string.Join("\r\n", errors.Select(e => string.Format("{0} in {1}", e.Description, e.Path.XPath)).ToArray());
                throw new ApplicationException(message);
            }
            return true;
        }

        private void CreatePackage(String filePath, SpreadsheetDocumentType spreadsheetDocumentType, bool validate)
        {
            PathHelper.CreateDirectory(Path.GetDirectoryName(filePath));
            var package = File.Exists(filePath)
                ? SpreadsheetDocument.Open(filePath, true)
                : SpreadsheetDocument.Create(filePath, spreadsheetDocumentType);

            using (package)
            {
                CreateParts(package);
                if (validate) Validate(package);
            }
        }

        private void CreatePackage(Stream stream, bool newStream, SpreadsheetDocumentType spreadsheetDocumentType, bool validate)
        {
            var package = newStream
                ? SpreadsheetDocument.Create(stream, spreadsheetDocumentType)
                : SpreadsheetDocument.Open(stream, true);

            using (package)
            {
                CreateParts(package);
                if (validate) Validate(package);
            }
        }

        // http://blogs.msdn.com/b/vsod/archive/2010/02/05/how-to-delete-a-worksheet-from-excel-using-open-xml-sdk-2-0.aspx
        private void DeleteSheetAndDependencies(WorkbookPart wbPart, string sheetId)
        {
            //Get the SheetToDelete from workbook.xml
            Sheet worksheet = wbPart.Workbook.Descendants<Sheet>().Where(s => s.Id == sheetId).FirstOrDefault();
            if (worksheet == null)
            { }

            string sheetName = worksheet.Name;
            // Get the pivot Table Parts
            IEnumerable<PivotTableCacheDefinitionPart> pvtTableCacheParts = wbPart.PivotTableCacheDefinitionParts;
            Dictionary<PivotTableCacheDefinitionPart, string> pvtTableCacheDefinationPart = new Dictionary<PivotTableCacheDefinitionPart, string>();
            foreach (PivotTableCacheDefinitionPart Item in pvtTableCacheParts)
            {
                PivotCacheDefinition pvtCacheDef = Item.PivotCacheDefinition;
                //Check if this CacheSource is linked to SheetToDelete
                var pvtCahce = pvtCacheDef.Descendants<CacheSource>().Where(s => s.WorksheetSource.Sheet == sheetName);
                if (pvtCahce.Count() > 0)
                {
                    pvtTableCacheDefinationPart.Add(Item, Item.ToString());
                }
            }
            foreach (var Item in pvtTableCacheDefinationPart)
            {
                wbPart.DeletePart(Item.Key);
            }

            // Remove the sheet reference from the workbook.
            WorksheetPart worksheetPart = (WorksheetPart)(wbPart.GetPartById(sheetId));
            worksheet.Remove();

            // Delete the worksheet part.
            wbPart.DeletePart(worksheetPart);

            //Get the DefinedNames
            var definedNames = wbPart.Workbook.Descendants<DefinedNames>().FirstOrDefault();
            if (definedNames != null)
            {
                List<DefinedName> defNamesToDelete = new List<DefinedName>();

                foreach (DefinedName Item in definedNames)
                {
                    // This condition checks to delete only those names which are part of Sheet in question
                    if (Item.Text.Contains(worksheet.Name + "!"))
                        defNamesToDelete.Add(Item);
                }

                foreach (DefinedName Item in defNamesToDelete)
                {
                    Item.Remove();
                }

            }
            // Get the CalculationChainPart
            //Note: An instance of this part type contains an ordered set of references to all cells in all worksheets in the
            //workbook whose value is calculated from any formula

            CalculationChainPart calChainPart;
            calChainPart = wbPart.CalculationChainPart;
            if (calChainPart != null)
            {
                var calChainEntries = calChainPart.CalculationChain.Descendants<CalculationCell>().Where(c => c.SheetId == sheetId);
                List<CalculationCell> calcsToDelete = new List<CalculationCell>();
                foreach (CalculationCell Item in calChainEntries)
                {
                    calcsToDelete.Add(Item);
                }

                foreach (CalculationCell Item in calcsToDelete)
                {
                    Item.Remove();
                }

                if (calChainPart.CalculationChain.Count() == 0)
                {
                    wbPart.DeletePart(calChainPart);
                }
            }
        }

        // Adds child parts and generates content of the specified part.
        private void CreateParts(SpreadsheetDocument document)
        {
            var context = new SaveContext();

            var workbookPart = document.WorkbookPart ?? document.AddWorkbookPart();

            var worksheets = WorksheetsInternal;


            var partsToRemove = workbookPart.Parts.Where(s => worksheets.Deleted.Contains(s.RelationshipId)).ToList();

            var pivotCacheDefinitionsToRemove = partsToRemove.SelectMany(s => ((WorksheetPart)s.OpenXmlPart).PivotTableParts.Select(pt => pt.PivotTableCacheDefinitionPart)).Distinct().ToList();
            pivotCacheDefinitionsToRemove.ForEach(c => workbookPart.DeletePart(c));

            if (workbookPart.Workbook != null && workbookPart.Workbook.PivotCaches != null)
            {
                var pivotCachesToRemove = workbookPart.Workbook.PivotCaches.Where(pc => pivotCacheDefinitionsToRemove.Select(pcd => workbookPart.GetIdOfPart(pcd)).ToList().Contains(((PivotCache)pc).Id)).Distinct().ToList();
                pivotCachesToRemove.ForEach(c => workbookPart.Workbook.PivotCaches.RemoveChild(c));
            }

            worksheets.Deleted.ToList().ForEach(ws => DeleteSheetAndDependencies(workbookPart, ws));

            // Ensure all RelId's have been added to the context
            context.RelIdGenerator.AddValues(workbookPart.Parts.Select(p => p.RelationshipId), RelType.Workbook);
            context.RelIdGenerator.AddValues(WorksheetsInternal.Cast<XLWorksheet>().Where(ws => !XLHelper.IsNullOrWhiteSpace(ws.RelId)).Select(ws => ws.RelId), RelType.Workbook);
            context.RelIdGenerator.AddValues(WorksheetsInternal.Cast<XLWorksheet>().Where(ws => !XLHelper.IsNullOrWhiteSpace(ws.LegacyDrawingId)).Select(ws => ws.LegacyDrawingId), RelType.Workbook);
            context.RelIdGenerator.AddValues(WorksheetsInternal
                .Cast<XLWorksheet>()
                .SelectMany(ws => ws.Tables.Cast<XLTable>())
                .Where(t => !XLHelper.IsNullOrWhiteSpace(t.RelId))
                .Select(t => t.RelId), RelType.Workbook);

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

            foreach (var worksheet in WorksheetsInternal.Cast<XLWorksheet>().OrderBy(w => w.Position))
            {
                //context.RelIdGenerator.Reset(RelType.);
                WorksheetPart worksheetPart;
                var wsRelId = worksheet.RelId;
                if (workbookPart.Parts.Any(p => p.RelationshipId == wsRelId))
                {
                    worksheetPart = (WorksheetPart)workbookPart.GetPartById(wsRelId);
                    var wsPartsToRemove = worksheetPart.TableDefinitionParts.ToList();
                    wsPartsToRemove.ForEach(tdp => worksheetPart.DeletePart(tdp));
                }
                else
                    worksheetPart = workbookPart.AddNewPart<WorksheetPart>(wsRelId);


                context.RelIdGenerator.AddValues(worksheetPart.HyperlinkRelationships.Select(hr => hr.Id), RelType.Workbook);
                context.RelIdGenerator.AddValues(worksheetPart.Parts.Select(p => p.RelationshipId), RelType.Workbook);
                if (worksheetPart.DrawingsPart != null)
                    context.RelIdGenerator.AddValues(worksheetPart.DrawingsPart.Parts.Select(p => p.RelationshipId), RelType.Workbook);

                // delete comment related parts (todo: review)
                DeleteComments(worksheetPart, worksheet, context);

                if (worksheet.Internals.CellsCollection.GetCells(c => c.HasComment).Any())
                {
                    var id = context.RelIdGenerator.GetNext(RelType.Workbook);
                    var worksheetCommentsPart =
                        worksheetPart.AddNewPart<WorksheetCommentsPart>(id);

                    GenerateWorksheetCommentsPartContent(worksheetCommentsPart, worksheet);

                    //VmlDrawingPart vmlDrawingPart = worksheetPart.AddNewPart<VmlDrawingPart>(worksheet.LegacyDrawingId);
                    var vmlDrawingPart = worksheetPart.VmlDrawingParts.FirstOrDefault();
                    if (vmlDrawingPart == null)
                    {
                        if (XLHelper.IsNullOrWhiteSpace(worksheet.LegacyDrawingId))
                        {
                            worksheet.LegacyDrawingId = context.RelIdGenerator.GetNext(RelType.Workbook);
                            worksheet.LegacyDrawingIsNew = true;
                        }

                        vmlDrawingPart = worksheetPart.AddNewPart<VmlDrawingPart>(worksheet.LegacyDrawingId);
                    }
                    GenerateVmlDrawingPartContent(vmlDrawingPart, worksheet, context);
                }

                GenerateWorksheetPartContent(worksheetPart, worksheet, context);

                if (worksheet.PivotTables.Any())
                {
                    GeneratePivotTables(workbookPart, worksheetPart, worksheet, context);
                }

                // Remove any orphaned references - maybe more types?
                foreach (var orphan in worksheetPart.Worksheet.OfType<LegacyDrawing>().Where(lg => !worksheetPart.Parts.Any(p => p.RelationshipId == lg.Id)))
                    worksheetPart.Worksheet.RemoveChild(orphan);

                foreach (var orphan in worksheetPart.Worksheet.OfType<Drawing>().Where(d => !worksheetPart.Parts.Any(p => p.RelationshipId == d.Id)))
                    worksheetPart.Worksheet.RemoveChild(orphan);
            }

            // Remove empty pivot cache part
            if (workbookPart.Workbook.PivotCaches != null && !workbookPart.Workbook.PivotCaches.Any())
                workbookPart.Workbook.RemoveChild(workbookPart.Workbook.PivotCaches);

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

        private void DeleteComments(WorksheetPart worksheetPart, XLWorksheet worksheet, SaveContext context)
        {
            // We have the comments so we can delete the comments part
            worksheetPart.DeletePart(worksheetPart.WorksheetCommentsPart);
            var vmlDrawingPart = worksheetPart.VmlDrawingParts.FirstOrDefault();

            // Only delete the VmlDrawingParts for comments.
            if (vmlDrawingPart != null)
            {
                var xdoc = XDocumentExtensions.Load(vmlDrawingPart.GetStream(FileMode.Open));
                //xdoc.Root.Elements().Where(e => e.Name.LocalName == "shapelayout").Remove();
                xdoc.Root.Elements().Where(
                    e => e.Name.LocalName == "shapetype" && (string)e.Attribute("id") == @"_x0000_t202").Remove();
                xdoc.Root.Elements().Where(
                    e => e.Name.LocalName == "shape" && (string)e.Attribute("type") == @"#_x0000_t202").Remove();
                var imageParts = vmlDrawingPart.ImageParts.ToList();
                var legacyParts = vmlDrawingPart.LegacyDiagramTextParts.ToList();
                var rId = worksheetPart.GetIdOfPart(vmlDrawingPart);
                worksheet.LegacyDrawingId = rId;
                worksheetPart.ChangeIdOfPart(vmlDrawingPart, "xxRRxx"); // Anything will do for the new relationship id
                // we just want it alive enough to create the copy

                var hasShapes = xdoc.Root.Elements().Any(e => e.Name.LocalName == "shape" || e.Name.LocalName == "group");

                VmlDrawingPart vmlDrawingPartNew = null;
                var hasNewPart = (imageParts.Count > 0 || legacyParts.Count > 0 || hasShapes);
                if (hasNewPart)
                {
                    vmlDrawingPartNew = worksheetPart.AddNewPart<VmlDrawingPart>(rId);

                    using (var writer = new XmlTextWriter(vmlDrawingPartNew.GetStream(FileMode.Create), Encoding.UTF8))
                    {
                        writer.WriteRaw(xdoc.ToString());
                    }

                    imageParts.ForEach(p => vmlDrawingPartNew.AddPart(p, vmlDrawingPart.GetIdOfPart(p)));
                    legacyParts.ForEach(p => vmlDrawingPartNew.AddPart(p, vmlDrawingPart.GetIdOfPart(p)));
                }

                worksheetPart.DeletePart(vmlDrawingPart);

                if (hasNewPart && rId != worksheetPart.GetIdOfPart(vmlDrawingPartNew))
                    worksheetPart.ChangeIdOfPart(vmlDrawingPartNew, rId);
            }
        }

        private static void GenerateTables(XLWorksheet worksheet, WorksheetPart worksheetPart, SaveContext context)
        {
            worksheetPart.Worksheet.RemoveAllChildren<TablePart>();

            if (!worksheet.Tables.Any()) return;

            foreach (var table in worksheet.Tables)
            {
                var tableRelId = context.RelIdGenerator.GetNext(RelType.Workbook);

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
                properties.AppendChild(new Application { Text = "Microsoft Excel" });

            if (properties.DocumentSecurity == null)
                properties.AppendChild(new DocumentSecurity { Text = "0" });

            if (properties.ScaleCrop == null)
                properties.AppendChild(new ScaleCrop { Text = "false" });

            if (properties.HeadingPairs == null)
                properties.HeadingPairs = new HeadingPairs();

            if (properties.TitlesOfParts == null)
                properties.TitlesOfParts = new TitlesOfParts();

            properties.HeadingPairs.VTVector = new VTVector { BaseType = VectorBaseValues.Variant };

            properties.TitlesOfParts.VTVector = new VTVector { BaseType = VectorBaseValues.Lpstr };

            var vTVectorOne = properties.HeadingPairs.VTVector;

            var vTVectorTwo = properties.TitlesOfParts.VTVector;

            var modifiedWorksheets =
                ((IEnumerable<XLWorksheet>)WorksheetsInternal).Select(w => new { w.Name, Order = w.Position }).ToList();
            var modifiedNamedRanges = GetModifiedNamedRanges();
            var modifiedWorksheetsCount = modifiedWorksheets.Count;
            var modifiedNamedRangesCount = modifiedNamedRanges.Count;

            InsertOnVtVector(vTVectorOne, "Worksheets", 0, modifiedWorksheetsCount.ToString());
            InsertOnVtVector(vTVectorOne, "Named Ranges", 2, modifiedNamedRangesCount.ToString());

            vTVectorTwo.Size = (UInt32)(modifiedNamedRangesCount + modifiedWorksheetsCount);

            foreach (
                var vTlpstr3 in modifiedWorksheets.OrderBy(w => w.Order).Select(w => new VTLPSTR { Text = w.Name }))
                vTVectorTwo.AppendChild(vTlpstr3);

            foreach (var vTlpstr7 in modifiedNamedRanges.Select(nr => new VTLPSTR { Text = nr }))
                vTVectorTwo.AppendChild(vTlpstr7);

            if (Properties.Manager != null)
            {
                if (!XLHelper.IsNullOrWhiteSpace(Properties.Manager))
                {
                    if (properties.Manager == null)
                        properties.Manager = new Manager();

                    properties.Manager.Text = Properties.Manager;
                }
                else
                    properties.Manager = null;
            }

            if (Properties.Company == null) return;

            if (!XLHelper.IsNullOrWhiteSpace(Properties.Company))
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
                var vTlpstr1 = new VTLPSTR { Text = property };
                variant1.AppendChild(vTlpstr1);
                vTVector.InsertAt(variant1, index);

                var variant2 = new Variant();
                var vTInt321 = new VTInt32();
                variant2.AppendChild(vTInt321);
                vTVector.InsertAt(variant2, index + 1);
            }

            var targetIndex = 0;
            foreach (var e in vTVector.Elements<Variant>())
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
            foreach (var w in WorksheetsInternal)
            {
                var wName = w.Name;
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

            if (Use1904DateSystem)
                workbook.WorkbookProperties.Date1904 = true;

            #endregion

            if (LockStructure || LockWindows)
            {
                if (workbook.WorkbookProtection == null)
                    workbook.WorkbookProtection = new WorkbookProtection();

                workbook.WorkbookProtection.LockStructure = LockStructure;
                workbook.WorkbookProtection.LockWindows = LockWindows;
            }
            else
            {
                workbook.WorkbookProtection = null;
            }


            if (workbook.BookViews == null)
                workbook.BookViews = new BookViews();

            if (workbook.Sheets == null)
                workbook.Sheets = new Sheets();

            var worksheets = WorksheetsInternal;
            workbook.Sheets.Elements<Sheet>().Where(s => worksheets.Deleted.Contains(s.Id)).ToList().ForEach(
                s => s.Remove());

            foreach (var sheet in workbook.Sheets.Elements<Sheet>())
            {
                var sheetId = (Int32)sheet.SheetId.Value;

                if (WorksheetsInternal.All<XLWorksheet>(w => w.SheetId != sheetId)) continue;

                var wks = WorksheetsInternal.Single<XLWorksheet>(w => w.SheetId == sheetId);
                wks.RelId = sheet.Id;
                sheet.Name = wks.Name;
            }

            foreach (var xlSheet in WorksheetsInternal.Cast<XLWorksheet>().OrderBy(w => w.Position))
            {
                string rId;
                if (xlSheet.SheetId == 0 && XLHelper.IsNullOrWhiteSpace(xlSheet.RelId))
                {
                    rId = context.RelIdGenerator.GetNext(RelType.Workbook);

                    while (WorksheetsInternal.Cast<XLWorksheet>().Any(w => w.SheetId == Int32.Parse(rId.Substring(3))))
                        rId = context.RelIdGenerator.GetNext(RelType.Workbook);

                    xlSheet.SheetId = Int32.Parse(rId.Substring(3));
                    xlSheet.RelId = rId;
                }
                else
                {
                    if (XLHelper.IsNullOrWhiteSpace(xlSheet.RelId))
                    {
                    rId = String.Format("rId{0}", xlSheet.SheetId);
                        context.RelIdGenerator.AddValues(new List<String> { rId }, RelType.Workbook);
                }
                    else
                        rId = xlSheet.RelId;
                }

                if (!workbook.Sheets.Cast<Sheet>().Any(s => s.Id == rId))
                {
                    var newSheet = new Sheet
                    {
                        Name = xlSheet.Name,
                        Id = rId,
                        SheetId = (UInt32)xlSheet.SheetId
                    };

                    workbook.Sheets.AppendChild(newSheet);
                }
            }

            var sheetElements = from sheet in workbook.Sheets.Elements<Sheet>()
                                join worksheet in ((IEnumerable<XLWorksheet>)WorksheetsInternal) on sheet.Id.Value
                                    equals worksheet.RelId
                                orderby worksheet.Position
                                select sheet;

            UInt32 firstSheetVisible = 0;
            var activeTab =
                (from us in UnsupportedSheets where us.IsActive select (UInt32)us.Position - 1).FirstOrDefault();
            var foundVisible = false;

            var totalSheets = sheetElements.Count() + UnsupportedSheets.Count;
            for (var p = 1; p <= totalSheets; p++)
            {
                if (UnsupportedSheets.All(us => us.Position != p))
                {
                    var sheet = sheetElements.ElementAt(p - UnsupportedSheets.Count(us => us.Position <= p) - 1);
                    workbook.Sheets.RemoveChild(sheet);
                    workbook.Sheets.AppendChild(sheet);
                    var xlSheet = Worksheet(sheet.Name);
                    if (xlSheet.Visibility != XLWorksheetVisibility.Visible)
                        sheet.State = xlSheet.Visibility.ToOpenXml();

                    if (foundVisible) continue;

                    if (sheet.State == null || sheet.State == SheetStateValues.Visible)
                        foundVisible = true;
                    else
                        firstSheetVisible++;
                }
                else
                {
                    var sheetId = UnsupportedSheets.First(us => us.Position == p).SheetId;
                    var sheet = workbook.Sheets.Elements<Sheet>().First(s => s.SheetId == sheetId);
                    workbook.Sheets.RemoveChild(sheet);
                    workbook.Sheets.AppendChild(sheet);
                }
            }

            var workbookView = workbook.BookViews.Elements<WorkbookView>().FirstOrDefault();

            if (activeTab == 0)
            {
                activeTab = firstSheetVisible;
                foreach (var ws in worksheets)
                {
                    if (!ws.TabActive) continue;

                    activeTab = (UInt32)(ws.Position - 1);
                    break;
                }
            }

            if (workbookView == null)
            {
                workbookView = new WorkbookView { ActiveTab = activeTab, FirstSheet = firstSheetVisible };
                workbook.BookViews.AppendChild(workbookView);
            }
            else
            {
                workbookView.ActiveTab = activeTab;
                workbookView.FirstSheet = firstSheetVisible;
            }

            var definedNames = new DefinedNames();
            foreach (var worksheet in WorksheetsInternal)
            {
                var wsSheetId = (UInt32)worksheet.SheetId;
                UInt32 sheetId = 0;
                foreach (var s in workbook.Sheets.Elements<Sheet>().TakeWhile(s => s.SheetId != wsSheetId))
                {
                    sheetId++;
                }

                if (worksheet.PageSetup.PrintAreas.Any())
                {
                    var definedName = new DefinedName { Name = "_xlnm.Print_Area", LocalSheetId = sheetId };
                    var worksheetName = worksheet.Name;
                    var definedNameText = worksheet.PageSetup.PrintAreas.Aggregate(String.Empty,
                        (current, printArea) =>
                            current +
                            ("'" + worksheetName + "'!" +
                             printArea.RangeAddress.
                                 FirstAddress.ToStringFixed(
                                     XLReferenceStyle.A1) +
                             ":" +
                             printArea.RangeAddress.
                                 LastAddress.ToStringFixed(
                                     XLReferenceStyle.A1) +
                             ","));
                    definedName.Text = definedNameText.Substring(0, definedNameText.Length - 1);
                    definedNames.AppendChild(definedName);
                }

                if (worksheet.AutoFilter.Enabled)
                {
                    var definedName = new DefinedName
                    {
                        Name = "_xlnm._FilterDatabase",
                        LocalSheetId = sheetId,
                        Text = "'" + worksheet.Name + "'!" +
                               worksheet.AutoFilter.Range.RangeAddress.FirstAddress.ToStringFixed(
                                   XLReferenceStyle.A1) +
                               ":" +
                               worksheet.AutoFilter.Range.RangeAddress.LastAddress.ToStringFixed(
                                   XLReferenceStyle.A1),
                        Hidden = BooleanValue.FromBoolean(true)
                    };
                    definedNames.AppendChild(definedName);
                }

                foreach (var nr in worksheet.NamedRanges.Where(n => n.Name != "_xlnm._FilterDatabase"))
                {
                    var definedName = new DefinedName
                    {
                        Name = nr.Name,
                        LocalSheetId = sheetId,
                        Text = nr.ToString()
                    };

                    if (!nr.Visible)
                        definedName.Hidden = BooleanValue.FromBoolean(true);

                    if (!XLHelper.IsNullOrWhiteSpace(nr.Comment))
                        definedName.Comment = nr.Comment;
                    definedNames.AppendChild(definedName);
                }


                var definedNameTextRow = String.Empty;
                var definedNameTextColumn = String.Empty;
                if (worksheet.PageSetup.FirstRowToRepeatAtTop > 0)
                {
                    definedNameTextRow = "'" + worksheet.Name + "'!" + worksheet.PageSetup.FirstRowToRepeatAtTop
                                         + ":" + worksheet.PageSetup.LastRowToRepeatAtTop;
                }
                if (worksheet.PageSetup.FirstColumnToRepeatAtLeft > 0)
                {
                    var minColumn = worksheet.PageSetup.FirstColumnToRepeatAtLeft;
                    var maxColumn = worksheet.PageSetup.LastColumnToRepeatAtLeft;
                    definedNameTextColumn = "'" + worksheet.Name + "'!" +
                                            XLHelper.GetColumnLetterFromNumber(minColumn)
                                            + ":" + XLHelper.GetColumnLetterFromNumber(maxColumn);
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

            foreach (var nr in NamedRanges)
            {
                var definedName = new DefinedName
                {
                    Name = nr.Name,
                    Text = nr.ToString()
                };

                if (!nr.Visible)
                    definedName.Hidden = BooleanValue.FromBoolean(true);

                if (!XLHelper.IsNullOrWhiteSpace(nr.Comment))
                    definedName.Comment = nr.Comment;
                definedNames.AppendChild(definedName);
            }

            workbook.DefinedNames = definedNames;

            if (workbook.CalculationProperties == null)
                workbook.CalculationProperties = new CalculationProperties { CalculationId = 125725U };

            if (CalculateMode == XLCalculateMode.Default)
                workbook.CalculationProperties.CalculationMode = null;
            else
                workbook.CalculationProperties.CalculationMode = CalculateMode.ToOpenXml();

            if (ReferenceStyle == XLReferenceStyle.Default)
                workbook.CalculationProperties.ReferenceMode = null;
            else
                workbook.CalculationProperties.ReferenceMode = ReferenceStyle.ToOpenXml();

            if (CalculationOnSave) workbook.CalculationProperties.CalculationOnSave = CalculationOnSave;
            if (ForceFullCalculation) workbook.CalculationProperties.ForceFullCalculation = ForceFullCalculation;
            if (FullCalculationOnLoad) workbook.CalculationProperties.FullCalculationOnLoad = FullCalculationOnLoad;
            if (FullPrecision) workbook.CalculationProperties.FullPrecision = FullPrecision;
        }

        private void GenerateSharedStringTablePartContent(SharedStringTablePart sharedStringTablePart,
            SaveContext context)
        {
            // Call all table headers to make sure their names are filled
            var x = 0;
            Worksheets.ForEach(w => w.Tables.ForEach(t => x = (t as XLTable).FieldNames.Count));

            sharedStringTablePart.SharedStringTable = new SharedStringTable { Count = 0, UniqueCount = 0 };

            var stringId = 0;

            var newStrings = new Dictionary<String, Int32>();
            var newRichStrings = new Dictionary<IXLRichText, Int32>();
            foreach (
                var c in
                    Worksheets.Cast<XLWorksheet>().SelectMany(
                        w =>
                            w.Internals.CellsCollection.GetCells(
                                c => ((c.DataType == XLCellValues.Text && c.ShareString) || c.HasRichText)
                                     && (c as XLCell).InnerText.Length > 0
                                     && XLHelper.IsNullOrWhiteSpace(c.FormulaA1)
                                )))
            {
                c.DataType = XLCellValues.Text;
                if (c.HasRichText)
                {
                    if (newRichStrings.ContainsKey(c.RichText))
                        c.SharedStringId = newRichStrings[c.RichText];
                    else
                    {
                        var sharedStringItem = new SharedStringItem();
                        foreach (var rt in c.RichText.Where(r => !String.IsNullOrEmpty(r.Text)))
                        {
                            sharedStringItem.Append(GetRun(rt));
                        }

                        if (c.RichText.HasPhonetics)
                        {
                            foreach (var p in c.RichText.Phonetics)
                            {
                                var phoneticRun = new PhoneticRun
                                {
                                    BaseTextStartIndex = (UInt32)p.Start,
                                    EndingBaseIndex = (UInt32)p.End
                                };

                                var text = new Text { Text = p.Text };
                                if (p.Text.PreserveSpaces())
                                    text.Space = SpaceProcessingModeValues.Preserve;

                                phoneticRun.Append(text);
                                sharedStringItem.Append(phoneticRun);
                            }
                            var f = new XLFont(null, c.RichText.Phonetics);
                            if (!context.SharedFonts.ContainsKey(f))
                                context.SharedFonts.Add(f, new FontInfo { Font = f });

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
                        var s = c.Value.ToString();
                        var sharedStringItem = new SharedStringItem();
                        var text = new Text { Text = XmlEncoder.EncodeString(s) };
                        if (!s.Trim().Equals(s))
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

        private static Run GetRun(IXLRichString rt)
        {
            var run = new Run();

            var runProperties = new RunProperties();

            var bold = rt.Bold ? new Bold() : null;
            var italic = rt.Italic ? new Italic() : null;
            var underline = rt.Underline != XLFontUnderlineValues.None
                ? new Underline { Val = rt.Underline.ToOpenXml() }
                : null;
            var strike = rt.Strikethrough ? new Strike() : null;
            var verticalAlignment = new VerticalTextAlignment
            { Val = rt.VerticalAlignment.ToOpenXml() };
            var shadow = rt.Shadow ? new Shadow() : null;
            var fontSize = new FontSize { Val = rt.FontSize };
            var color = GetNewColor(rt.FontColor);
            var fontName = new RunFont { Val = rt.FontName };
            var fontFamilyNumbering = new FontFamily { Val = (Int32)rt.FontFamilyNumbering };

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

            var text = new Text { Text = rt.Text };
            if (rt.Text.PreserveSpaces())
                text.Space = SpaceProcessingModeValues.Preserve;

            run.Append(runProperties);
            run.Append(text);
            return run;
        }

        private void GenerateCalculationChainPartContent(WorkbookPart workbookPart, SaveContext context)
        {
            var thisRelId = context.RelIdGenerator.GetNext(RelType.Workbook);
            if (workbookPart.CalculationChainPart == null)
                workbookPart.AddNewPart<CalculationChainPart>(thisRelId);

            if (workbookPart.CalculationChainPart.CalculationChain == null)
                workbookPart.CalculationChainPart.CalculationChain = new CalculationChain();

            var calculationChain = workbookPart.CalculationChainPart.CalculationChain;
            calculationChain.RemoveAllChildren<CalculationCell>();

            foreach (var worksheet in WorksheetsInternal)
            {
                var cellsWithoutFormulas = new HashSet<String>();
                foreach (var c in worksheet.Internals.CellsCollection.GetCells())
                {
                    if (XLHelper.IsNullOrWhiteSpace(c.FormulaA1))
                        cellsWithoutFormulas.Add(c.Address.ToStringRelative());
                    else
                    {
                        if (c.FormulaA1.StartsWith("{"))
                        {
                            var cc = new CalculationCell
                            {
                                CellReference = c.Address.ToString(),
                                SheetId = worksheet.SheetId
                            };

                            if (c.FormulaReference == null)
                                c.FormulaReference = c.AsRange().RangeAddress;
                            if (c.FormulaReference.FirstAddress.Equals(c.Address))
                            {
                                cc.Array = true;
                                calculationChain.AppendChild(cc);
                                calculationChain.AppendChild(new CalculationCell { CellReference = c.Address.ToString(), InChildChain = true });
                            }
                            else
                            {
                                calculationChain.AppendChild(cc);
                            }
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
            var theme1 = new Theme { Name = "Office Theme" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            var themeElements1 = new ThemeElements();

            var colorScheme1 = new ColorScheme { Name = "Office" };

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
            var rgbColorModelHex1 = new RgbColorModelHex { Val = Theme.Text2.Color.ToHex().Substring(2) };

            dark2Color1.AppendChild(rgbColorModelHex1);

            var light2Color1 = new Light2Color();
            var rgbColorModelHex2 = new RgbColorModelHex { Val = Theme.Background2.Color.ToHex().Substring(2) };

            light2Color1.AppendChild(rgbColorModelHex2);

            var accent1Color1 = new Accent1Color();
            var rgbColorModelHex3 = new RgbColorModelHex { Val = Theme.Accent1.Color.ToHex().Substring(2) };

            accent1Color1.AppendChild(rgbColorModelHex3);

            var accent2Color1 = new Accent2Color();
            var rgbColorModelHex4 = new RgbColorModelHex { Val = Theme.Accent2.Color.ToHex().Substring(2) };

            accent2Color1.AppendChild(rgbColorModelHex4);

            var accent3Color1 = new Accent3Color();
            var rgbColorModelHex5 = new RgbColorModelHex { Val = Theme.Accent3.Color.ToHex().Substring(2) };

            accent3Color1.AppendChild(rgbColorModelHex5);

            var accent4Color1 = new Accent4Color();
            var rgbColorModelHex6 = new RgbColorModelHex { Val = Theme.Accent4.Color.ToHex().Substring(2) };

            accent4Color1.AppendChild(rgbColorModelHex6);

            var accent5Color1 = new Accent5Color();
            var rgbColorModelHex7 = new RgbColorModelHex { Val = Theme.Accent5.Color.ToHex().Substring(2) };

            accent5Color1.AppendChild(rgbColorModelHex7);

            var accent6Color1 = new Accent6Color();
            var rgbColorModelHex8 = new RgbColorModelHex { Val = Theme.Accent6.Color.ToHex().Substring(2) };

            accent6Color1.AppendChild(rgbColorModelHex8);

            var hyperlink1 = new DocumentFormat.OpenXml.Drawing.Hyperlink();
            var rgbColorModelHex9 = new RgbColorModelHex { Val = Theme.Hyperlink.Color.ToHex().Substring(2) };

            hyperlink1.AppendChild(rgbColorModelHex9);

            var followedHyperlinkColor1 = new FollowedHyperlinkColor();
            var rgbColorModelHex10 = new RgbColorModelHex { Val = Theme.FollowedHyperlink.Color.ToHex().Substring(2) };

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

            var fontScheme2 = new FontScheme { Name = "Office" };

            var majorFont1 = new MajorFont();
            var latinFont1 = new LatinFont { Typeface = "Cambria" };
            var eastAsianFont1 = new EastAsianFont { Typeface = "" };
            var complexScriptFont1 = new ComplexScriptFont { Typeface = "" };
            var supplementalFont1 = new SupplementalFont { Script = "Jpan", Typeface = "ＭＳ Ｐゴシック" };
            var supplementalFont2 = new SupplementalFont { Script = "Hang", Typeface = "맑은 고딕" };
            var supplementalFont3 = new SupplementalFont { Script = "Hans", Typeface = "宋体" };
            var supplementalFont4 = new SupplementalFont { Script = "Hant", Typeface = "新細明體" };
            var supplementalFont5 = new SupplementalFont { Script = "Arab", Typeface = "Times New Roman" };
            var supplementalFont6 = new SupplementalFont { Script = "Hebr", Typeface = "Times New Roman" };
            var supplementalFont7 = new SupplementalFont { Script = "Thai", Typeface = "Tahoma" };
            var supplementalFont8 = new SupplementalFont { Script = "Ethi", Typeface = "Nyala" };
            var supplementalFont9 = new SupplementalFont { Script = "Beng", Typeface = "Vrinda" };
            var supplementalFont10 = new SupplementalFont { Script = "Gujr", Typeface = "Shruti" };
            var supplementalFont11 = new SupplementalFont { Script = "Khmr", Typeface = "MoolBoran" };
            var supplementalFont12 = new SupplementalFont { Script = "Knda", Typeface = "Tunga" };
            var supplementalFont13 = new SupplementalFont { Script = "Guru", Typeface = "Raavi" };
            var supplementalFont14 = new SupplementalFont { Script = "Cans", Typeface = "Euphemia" };
            var supplementalFont15 = new SupplementalFont { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            var supplementalFont16 = new SupplementalFont { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            var supplementalFont17 = new SupplementalFont { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            var supplementalFont18 = new SupplementalFont { Script = "Thaa", Typeface = "MV Boli" };
            var supplementalFont19 = new SupplementalFont { Script = "Deva", Typeface = "Mangal" };
            var supplementalFont20 = new SupplementalFont { Script = "Telu", Typeface = "Gautami" };
            var supplementalFont21 = new SupplementalFont { Script = "Taml", Typeface = "Latha" };
            var supplementalFont22 = new SupplementalFont { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            var supplementalFont23 = new SupplementalFont { Script = "Orya", Typeface = "Kalinga" };
            var supplementalFont24 = new SupplementalFont { Script = "Mlym", Typeface = "Kartika" };
            var supplementalFont25 = new SupplementalFont { Script = "Laoo", Typeface = "DokChampa" };
            var supplementalFont26 = new SupplementalFont { Script = "Sinh", Typeface = "Iskoola Pota" };
            var supplementalFont27 = new SupplementalFont { Script = "Mong", Typeface = "Mongolian Baiti" };
            var supplementalFont28 = new SupplementalFont { Script = "Viet", Typeface = "Times New Roman" };
            var supplementalFont29 = new SupplementalFont { Script = "Uigh", Typeface = "Microsoft Uighur" };

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
            var latinFont2 = new LatinFont { Typeface = "Calibri" };
            var eastAsianFont2 = new EastAsianFont { Typeface = "" };
            var complexScriptFont2 = new ComplexScriptFont { Typeface = "" };
            var supplementalFont30 = new SupplementalFont { Script = "Jpan", Typeface = "ＭＳ Ｐゴシック" };
            var supplementalFont31 = new SupplementalFont { Script = "Hang", Typeface = "맑은 고딕" };
            var supplementalFont32 = new SupplementalFont { Script = "Hans", Typeface = "宋体" };
            var supplementalFont33 = new SupplementalFont { Script = "Hant", Typeface = "新細明體" };
            var supplementalFont34 = new SupplementalFont { Script = "Arab", Typeface = "Arial" };
            var supplementalFont35 = new SupplementalFont { Script = "Hebr", Typeface = "Arial" };
            var supplementalFont36 = new SupplementalFont { Script = "Thai", Typeface = "Tahoma" };
            var supplementalFont37 = new SupplementalFont { Script = "Ethi", Typeface = "Nyala" };
            var supplementalFont38 = new SupplementalFont { Script = "Beng", Typeface = "Vrinda" };
            var supplementalFont39 = new SupplementalFont { Script = "Gujr", Typeface = "Shruti" };
            var supplementalFont40 = new SupplementalFont { Script = "Khmr", Typeface = "DaunPenh" };
            var supplementalFont41 = new SupplementalFont { Script = "Knda", Typeface = "Tunga" };
            var supplementalFont42 = new SupplementalFont { Script = "Guru", Typeface = "Raavi" };
            var supplementalFont43 = new SupplementalFont { Script = "Cans", Typeface = "Euphemia" };
            var supplementalFont44 = new SupplementalFont { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            var supplementalFont45 = new SupplementalFont { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            var supplementalFont46 = new SupplementalFont { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            var supplementalFont47 = new SupplementalFont { Script = "Thaa", Typeface = "MV Boli" };
            var supplementalFont48 = new SupplementalFont { Script = "Deva", Typeface = "Mangal" };
            var supplementalFont49 = new SupplementalFont { Script = "Telu", Typeface = "Gautami" };
            var supplementalFont50 = new SupplementalFont { Script = "Taml", Typeface = "Latha" };
            var supplementalFont51 = new SupplementalFont { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            var supplementalFont52 = new SupplementalFont { Script = "Orya", Typeface = "Kalinga" };
            var supplementalFont53 = new SupplementalFont { Script = "Mlym", Typeface = "Kartika" };
            var supplementalFont54 = new SupplementalFont { Script = "Laoo", Typeface = "DokChampa" };
            var supplementalFont55 = new SupplementalFont { Script = "Sinh", Typeface = "Iskoola Pota" };
            var supplementalFont56 = new SupplementalFont { Script = "Mong", Typeface = "Mongolian Baiti" };
            var supplementalFont57 = new SupplementalFont { Script = "Viet", Typeface = "Arial" };
            var supplementalFont58 = new SupplementalFont { Script = "Uigh", Typeface = "Microsoft Uighur" };

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

            var formatScheme1 = new FormatScheme { Name = "Office" };

            var fillStyleList1 = new FillStyleList();

            var solidFill1 = new SolidFill();
            var schemeColor1 = new SchemeColor { Val = SchemeColorValues.PhColor };

            solidFill1.AppendChild(schemeColor1);

            var gradientFill1 = new GradientFill { RotateWithShape = true };

            var gradientStopList1 = new GradientStopList();

            var gradientStop1 = new GradientStop { Position = 0 };

            var schemeColor2 = new SchemeColor { Val = SchemeColorValues.PhColor };
            var tint1 = new Tint { Val = 50000 };
            var saturationModulation1 = new SaturationModulation { Val = 300000 };

            schemeColor2.AppendChild(tint1);
            schemeColor2.AppendChild(saturationModulation1);

            gradientStop1.AppendChild(schemeColor2);

            var gradientStop2 = new GradientStop { Position = 35000 };

            var schemeColor3 = new SchemeColor { Val = SchemeColorValues.PhColor };
            var tint2 = new Tint { Val = 37000 };
            var saturationModulation2 = new SaturationModulation { Val = 300000 };

            schemeColor3.AppendChild(tint2);
            schemeColor3.AppendChild(saturationModulation2);

            gradientStop2.AppendChild(schemeColor3);

            var gradientStop3 = new GradientStop { Position = 100000 };

            var schemeColor4 = new SchemeColor { Val = SchemeColorValues.PhColor };
            var tint3 = new Tint { Val = 15000 };
            var saturationModulation3 = new SaturationModulation { Val = 350000 };

            schemeColor4.AppendChild(tint3);
            schemeColor4.AppendChild(saturationModulation3);

            gradientStop3.AppendChild(schemeColor4);

            gradientStopList1.AppendChild(gradientStop1);
            gradientStopList1.AppendChild(gradientStop2);
            gradientStopList1.AppendChild(gradientStop3);
            var linearGradientFill1 = new LinearGradientFill { Angle = 16200000, Scaled = true };

            gradientFill1.AppendChild(gradientStopList1);
            gradientFill1.AppendChild(linearGradientFill1);

            var gradientFill2 = new GradientFill { RotateWithShape = true };

            var gradientStopList2 = new GradientStopList();

            var gradientStop4 = new GradientStop { Position = 0 };

            var schemeColor5 = new SchemeColor { Val = SchemeColorValues.PhColor };
            var shade1 = new Shade { Val = 51000 };
            var saturationModulation4 = new SaturationModulation { Val = 130000 };

            schemeColor5.AppendChild(shade1);
            schemeColor5.AppendChild(saturationModulation4);

            gradientStop4.AppendChild(schemeColor5);

            var gradientStop5 = new GradientStop { Position = 80000 };

            var schemeColor6 = new SchemeColor { Val = SchemeColorValues.PhColor };
            var shade2 = new Shade { Val = 93000 };
            var saturationModulation5 = new SaturationModulation { Val = 130000 };

            schemeColor6.AppendChild(shade2);
            schemeColor6.AppendChild(saturationModulation5);

            gradientStop5.AppendChild(schemeColor6);

            var gradientStop6 = new GradientStop { Position = 100000 };

            var schemeColor7 = new SchemeColor { Val = SchemeColorValues.PhColor };
            var shade3 = new Shade { Val = 94000 };
            var saturationModulation6 = new SaturationModulation { Val = 135000 };

            schemeColor7.AppendChild(shade3);
            schemeColor7.AppendChild(saturationModulation6);

            gradientStop6.AppendChild(schemeColor7);

            gradientStopList2.AppendChild(gradientStop4);
            gradientStopList2.AppendChild(gradientStop5);
            gradientStopList2.AppendChild(gradientStop6);
            var linearGradientFill2 = new LinearGradientFill { Angle = 16200000, Scaled = false };

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

            var schemeColor8 = new SchemeColor { Val = SchemeColorValues.PhColor };
            var shade4 = new Shade { Val = 95000 };
            var saturationModulation7 = new SaturationModulation { Val = 105000 };

            schemeColor8.AppendChild(shade4);
            schemeColor8.AppendChild(saturationModulation7);

            solidFill2.AppendChild(schemeColor8);
            var presetDash1 = new PresetDash { Val = PresetLineDashValues.Solid };

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
            var schemeColor9 = new SchemeColor { Val = SchemeColorValues.PhColor };

            solidFill3.AppendChild(schemeColor9);
            var presetDash2 = new PresetDash { Val = PresetLineDashValues.Solid };

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
            var schemeColor10 = new SchemeColor { Val = SchemeColorValues.PhColor };

            solidFill4.AppendChild(schemeColor10);
            var presetDash3 = new PresetDash { Val = PresetLineDashValues.Solid };

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

            var rgbColorModelHex11 = new RgbColorModelHex { Val = "000000" };
            var alpha1 = new Alpha { Val = 38000 };

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

            var rgbColorModelHex12 = new RgbColorModelHex { Val = "000000" };
            var alpha2 = new Alpha { Val = 35000 };

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

            var rgbColorModelHex13 = new RgbColorModelHex { Val = "000000" };
            var alpha3 = new Alpha { Val = 35000 };

            rgbColorModelHex13.AppendChild(alpha3);

            outerShadow3.AppendChild(rgbColorModelHex13);

            effectList3.AppendChild(outerShadow3);

            var scene3DType1 = new Scene3DType();

            var camera1 = new Camera { Preset = PresetCameraValues.OrthographicFront };
            var rotation1 = new Rotation { Latitude = 0, Longitude = 0, Revolution = 0 };

            camera1.AppendChild(rotation1);

            var lightRig1 = new LightRig { Rig = LightRigValues.ThreePoints, Direction = LightRigDirectionValues.Top };
            var rotation2 = new Rotation { Latitude = 0, Longitude = 0, Revolution = 1200000 };

            lightRig1.AppendChild(rotation2);

            scene3DType1.AppendChild(camera1);
            scene3DType1.AppendChild(lightRig1);

            var shape3DType1 = new Shape3DType();
            var bevelTop1 = new BevelTop { Width = 63500L, Height = 25400L };

            shape3DType1.AppendChild(bevelTop1);

            effectStyle3.AppendChild(effectList3);
            effectStyle3.AppendChild(scene3DType1);
            effectStyle3.AppendChild(shape3DType1);

            effectStyleList1.AppendChild(effectStyle1);
            effectStyleList1.AppendChild(effectStyle2);
            effectStyleList1.AppendChild(effectStyle3);

            var backgroundFillStyleList1 = new BackgroundFillStyleList();

            var solidFill5 = new SolidFill();
            var schemeColor11 = new SchemeColor { Val = SchemeColorValues.PhColor };

            solidFill5.AppendChild(schemeColor11);

            var gradientFill3 = new GradientFill { RotateWithShape = true };

            var gradientStopList3 = new GradientStopList();

            var gradientStop7 = new GradientStop { Position = 0 };

            var schemeColor12 = new SchemeColor { Val = SchemeColorValues.PhColor };
            var tint4 = new Tint { Val = 40000 };
            var saturationModulation8 = new SaturationModulation { Val = 350000 };

            schemeColor12.AppendChild(tint4);
            schemeColor12.AppendChild(saturationModulation8);

            gradientStop7.AppendChild(schemeColor12);

            var gradientStop8 = new GradientStop { Position = 40000 };

            var schemeColor13 = new SchemeColor { Val = SchemeColorValues.PhColor };
            var tint5 = new Tint { Val = 45000 };
            var shade5 = new Shade { Val = 99000 };
            var saturationModulation9 = new SaturationModulation { Val = 350000 };

            schemeColor13.AppendChild(tint5);
            schemeColor13.AppendChild(shade5);
            schemeColor13.AppendChild(saturationModulation9);

            gradientStop8.AppendChild(schemeColor13);

            var gradientStop9 = new GradientStop { Position = 100000 };

            var schemeColor14 = new SchemeColor { Val = SchemeColorValues.PhColor };
            var shade6 = new Shade { Val = 20000 };
            var saturationModulation10 = new SaturationModulation { Val = 255000 };

            schemeColor14.AppendChild(shade6);
            schemeColor14.AppendChild(saturationModulation10);

            gradientStop9.AppendChild(schemeColor14);

            gradientStopList3.AppendChild(gradientStop7);
            gradientStopList3.AppendChild(gradientStop8);
            gradientStopList3.AppendChild(gradientStop9);

            var pathGradientFill1 = new PathGradientFill { Path = PathShadeValues.Circle };
            var fillToRectangle1 = new FillToRectangle { Left = 50000, Top = -80000, Right = 50000, Bottom = 180000 };

            pathGradientFill1.AppendChild(fillToRectangle1);

            gradientFill3.AppendChild(gradientStopList3);
            gradientFill3.AppendChild(pathGradientFill1);

            var gradientFill4 = new GradientFill { RotateWithShape = true };

            var gradientStopList4 = new GradientStopList();

            var gradientStop10 = new GradientStop { Position = 0 };

            var schemeColor15 = new SchemeColor { Val = SchemeColorValues.PhColor };
            var tint6 = new Tint { Val = 80000 };
            var saturationModulation11 = new SaturationModulation { Val = 300000 };

            schemeColor15.AppendChild(tint6);
            schemeColor15.AppendChild(saturationModulation11);

            gradientStop10.AppendChild(schemeColor15);

            var gradientStop11 = new GradientStop { Position = 100000 };

            var schemeColor16 = new SchemeColor { Val = SchemeColorValues.PhColor };
            var shade7 = new Shade { Val = 30000 };
            var saturationModulation12 = new SaturationModulation { Val = 200000 };

            schemeColor16.AppendChild(shade7);
            schemeColor16.AppendChild(saturationModulation12);

            gradientStop11.AppendChild(schemeColor16);

            gradientStopList4.AppendChild(gradientStop10);
            gradientStopList4.AppendChild(gradientStop11);

            var pathGradientFill2 = new PathGradientFill { Path = PathShadeValues.Circle };
            var fillToRectangle2 = new FillToRectangle { Left = 50000, Top = 50000, Right = 50000, Bottom = 50000 };

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
            var propertyId = 1;
            foreach (var p in CustomProperties)
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
                    var vTlpwstr1 = new VTLPWSTR { Text = p.GetValue<string>() };
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
                        Text = p.GetValue<Double>().ToInvariantString()
                    };
                    customDocumentProperty.AppendChild(vTDouble1);
                }
                else
                {
                    var vTBool1 = new VTBool { Text = p.GetValue<Boolean>().ToString().ToLower() };
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
            var tableName = originalTableName.RemoveSpecialCharacters();
            var name = tableName;
            if (context.TableNames.Contains(name))
            {
                var i = 1;
                name = tableName + i.ToInvariantString();
                while (context.TableNames.Contains(name))
                {
                    i++;
                    name = tableName + i.ToInvariantString();
                }
            }

            context.TableNames.Add(name);
            return name;
        }

        private static void GenerateTableDefinitionPartContent(TableDefinitionPart tableDefinitionPart, XLTable xlTable,
            SaveContext context)
        {
            context.TableId++;
            var reference = xlTable.RangeAddress.FirstAddress + ":" + xlTable.RangeAddress.LastAddress;
            var tableName = GetTableName(xlTable.Name, context);
            var table = new Table
            {
                Id = context.TableId,
                Name = tableName,
                DisplayName = tableName,
                Reference = reference
            };

            if (!xlTable.ShowHeaderRow)
                table.HeaderRowCount = 0;

            if (xlTable.ShowTotalsRow)
                table.TotalsRowCount = 1;
            else
                table.TotalsRowShown = false;

            var tableColumns1 = new TableColumns { Count = (UInt32)xlTable.ColumnCount() };

            UInt32 columnId = 0;
            foreach (var xlField in xlTable.Fields)
            {
                columnId++;
                var fieldName = xlField.Name;
                var tableColumn1 = new TableColumn
                {
                    Id = columnId,
                    Name = fieldName.Replace("_x000a_", "_x005f_x000a_").Replace(Environment.NewLine, "_x000a_")
                };
                if (xlTable.ShowTotalsRow)
                {
                    if (xlField.TotalsRowFunction != XLTotalsRowFunction.None)
                    {
                        tableColumn1.TotalsRowFunction = xlField.TotalsRowFunction.ToOpenXml();

                        if (xlField.TotalsRowFunction == XLTotalsRowFunction.Custom)
                            tableColumn1.TotalsRowFormula = new TotalsRowFormula(xlField.TotalsRowFormulaA1);
                    }

                    if (!XLHelper.IsNullOrWhiteSpace(xlField.TotalsRowLabel))
                        tableColumn1.TotalsRowLabel = xlField.TotalsRowLabel;
                }
                tableColumns1.AppendChild(tableColumn1);
            }

            var tableStyleInfo1 = new TableStyleInfo
            {
                ShowFirstColumn = xlTable.EmphasizeFirstColumn,
                ShowLastColumn = xlTable.EmphasizeLastColumn,
                ShowRowStripes = xlTable.ShowRowStripes,
                ShowColumnStripes = xlTable.ShowColumnStripes
            };

            if (xlTable.Theme != XLTableTheme.None)
                tableStyleInfo1.Name = xlTable.Theme.Name;

            if (xlTable.ShowAutoFilter)
            {
                var autoFilter1 = new AutoFilter();
                if (xlTable.ShowTotalsRow)
                {
                    xlTable.AutoFilter.Range = xlTable.Worksheet.Range(
                        xlTable.RangeAddress.FirstAddress.RowNumber, xlTable.RangeAddress.FirstAddress.ColumnNumber,
                        xlTable.RangeAddress.LastAddress.RowNumber - 1, xlTable.RangeAddress.LastAddress.ColumnNumber);
                }
                else
                    xlTable.AutoFilter.Range = xlTable.Worksheet.Range(xlTable.RangeAddress);

                PopulateAutoFilter(xlTable.AutoFilter, autoFilter1);

                table.AppendChild(autoFilter1);
            }

            table.AppendChild(tableColumns1);
            table.AppendChild(tableStyleInfo1);

            tableDefinitionPart.Table = table;
        }

        private static void GeneratePivotTables(WorkbookPart workbookPart, WorksheetPart worksheetPart,
            XLWorksheet xlWorksheet,
            SaveContext context)
        {
            PivotCaches pivotCaches;
            uint cacheId = 0;
            if (workbookPart.Workbook.PivotCaches == null)
                pivotCaches = workbookPart.Workbook.InsertAfter(new PivotCaches(), workbookPart.Workbook.CalculationProperties);
            else
            {
                pivotCaches = workbookPart.Workbook.PivotCaches;
                if (pivotCaches.Any())
                    cacheId = pivotCaches.Cast<PivotCache>().Max(pc => pc.CacheId.Value) + 1;
            }

            foreach (var pt in xlWorksheet.PivotTables.Cast<XLPivotTable>())
            {
                // TODO: Avoid duplicate pivot caches of same source range

                var workbookCacheRelId = pt.WorkbookCacheRelId;
                PivotCache pivotCache;
                PivotTableCacheDefinitionPart pivotTableCacheDefinitionPart;
                if (!XLHelper.IsNullOrWhiteSpace(pt.WorkbookCacheRelId))
                {
                    pivotCache = pivotCaches.Cast<PivotCache>().Single(pc => pc.Id.Value == pt.WorkbookCacheRelId);
                    pivotTableCacheDefinitionPart = workbookPart.GetPartById(pt.WorkbookCacheRelId) as PivotTableCacheDefinitionPart;
                }
                else
                {
                    workbookCacheRelId = context.RelIdGenerator.GetNext(RelType.Workbook);
                    pivotCache = new PivotCache { CacheId = cacheId++, Id = workbookCacheRelId };
                    pivotTableCacheDefinitionPart = workbookPart.AddNewPart<PivotTableCacheDefinitionPart>(workbookCacheRelId);
                }

                GeneratePivotTableCacheDefinitionPartContent(pivotTableCacheDefinitionPart, pt);

                if (XLHelper.IsNullOrWhiteSpace(pt.WorkbookCacheRelId))
                    pivotCaches.AppendChild(pivotCache);

                PivotTablePart pivotTablePart;
                if (XLHelper.IsNullOrWhiteSpace(pt.RelId))
                    pivotTablePart = worksheetPart.AddNewPart<PivotTablePart>(context.RelIdGenerator.GetNext(RelType.Workbook));
                else
                    pivotTablePart = worksheetPart.GetPartById(pt.RelId) as PivotTablePart;

                GeneratePivotTablePartContent(pivotTablePart, pt, pivotCache.CacheId, context);

                if (XLHelper.IsNullOrWhiteSpace(pt.RelId))
                    pivotTablePart.AddPart(pivotTableCacheDefinitionPart, context.RelIdGenerator.GetNext(RelType.Workbook));
            }
        }

        // Generates content of pivotTableCacheDefinitionPart
        private static void GeneratePivotTableCacheDefinitionPartContent(
            PivotTableCacheDefinitionPart pivotTableCacheDefinitionPart, IXLPivotTable pt)
        {
            var source = pt.SourceRange;

            var pivotCacheDefinition = new PivotCacheDefinition
            {
                Id = "rId1",
                SaveData = pt.SaveSourceData,
                RefreshOnLoad = true //pt.RefreshDataOnOpen
            };
            if (pt.ItemsToRetainPerField == XLItemsToRetain.None)
                pivotCacheDefinition.MissingItemsLimit = 0U;
            else if (pt.ItemsToRetainPerField == XLItemsToRetain.Max)
                pivotCacheDefinition.MissingItemsLimit = XLHelper.MaxRowNumber;

            pivotCacheDefinition.AddNamespaceDeclaration("r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            var cacheSource = new CacheSource { Type = SourceValues.Worksheet };
            cacheSource.AppendChild(new WorksheetSource { Name = source.ToString() });

            var cacheFields = new CacheFields();

            foreach (var c in source.Columns())
            {
                var columnNumber = c.ColumnNumber();
                var columnName = c.FirstCell().Value.ToString();
                var xlpf = pt.Fields.Add(columnName);

                var field =
                    pt.RowLabels.Union(pt.ColumnLabels).Union(pt.ReportFilters).FirstOrDefault(f => f.SourceName == columnName);
                if (field != null)
                {
                    xlpf.CustomName = field.CustomName;
                    xlpf.Subtotals.AddRange(field.Subtotals);
                }

                var sharedItems = new SharedItems();

                var onlyNumbers =
                    !source.Cells().Any(
                        cell =>
                            cell.Address.ColumnNumber == columnNumber &&
                            cell.Address.RowNumber > source.FirstRow().RowNumber() && cell.DataType != XLCellValues.Number);
                if (onlyNumbers)
                {
                    sharedItems = new SharedItems
                    { ContainsSemiMixedTypes = false, ContainsString = false, ContainsNumber = true };
                }
                else
                {
                    foreach (var cellValue in source.Cells()
                        .Where(cell => cell.Address.ColumnNumber == columnNumber && cell.Address.RowNumber > source.FirstRow().RowNumber())
                        .Select(cell => cell.Value.ToString())
                        .Where(cellValue => !xlpf.SharedStrings.Select(ss => ss.ToLower()).Contains(cellValue.ToLower())))
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

            var pivotTableCacheRecordsPart = pivotTableCacheDefinitionPart.GetPartsOfType<PivotTableCacheRecordsPart>().Any() ?
                pivotTableCacheDefinitionPart.GetPartsOfType<PivotTableCacheRecordsPart>().First() :
                pivotTableCacheDefinitionPart.AddNewPart<PivotTableCacheRecordsPart>("rId1");

            var pivotCacheRecords = new PivotCacheRecords();
            pivotCacheRecords.AddNamespaceDeclaration("r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            pivotTableCacheRecordsPart.PivotCacheRecords = pivotCacheRecords;
        }

        // Generates content of pivotTablePart
        private static void GeneratePivotTablePartContent(PivotTablePart pivotTablePart, IXLPivotTable pt, uint cacheId, SaveContext context)
        {
            var pivotTableDefinition = new PivotTableDefinition
            {
                Name = pt.Name,
                CacheId = cacheId,
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
                GridDropZones = GetBooleanValue(pt.ClassicPivotTableLayout, false),
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

            var location = new Location
            {
                Reference = pt.TargetCell.Address.ToString(),
                FirstHeaderRow = 1U,
                FirstDataRow = 1U,
                FirstDataColumn = 1U
            };


            var rowFields = new RowFields();
            var columnFields = new ColumnFields();
            var rowItems = new RowItems();
            var columnItems = new ColumnItems();
            var pageFields = new PageFields { Count = (uint)pt.ReportFilters.Count() };
            var pivotFields = new PivotFields { Count = Convert.ToUInt32(pt.SourceRange.ColumnCount()) };

            foreach (var xlpf in pt.Fields.OrderBy(f => pt.RowLabels.Any(p => p.SourceName == f.SourceName) ? pt.RowLabels.IndexOf(f) : Int32.MaxValue))
            {
                if (pt.RowLabels.Any(p => p.SourceName == xlpf.SourceName))
                {
                    var f = new Field { Index = pt.Fields.IndexOf(xlpf) };
                    rowFields.AppendChild(f);

                    for (var i = 0; i < xlpf.SharedStrings.Count; i++)
                    {
                        var rowItem = new RowItem();
                        rowItem.AppendChild(new MemberPropertyIndex { Val = i });
                        rowItems.AppendChild(rowItem);
                    }

                    var rowItemTotal = new RowItem { ItemType = ItemValues.Grand };
                    rowItemTotal.AppendChild(new MemberPropertyIndex());
                    rowItems.AppendChild(rowItemTotal);
                }
                else if (pt.ColumnLabels.Any(p => p.SourceName == xlpf.SourceName))
                {
                    var f = new Field { Index = pt.Fields.IndexOf(xlpf) };
                    columnFields.AppendChild(f);

                    for (var i = 0; i < xlpf.SharedStrings.Count; i++)
                    {
                        var rowItem = new RowItem();
                        rowItem.AppendChild(new MemberPropertyIndex { Val = i });
                        columnItems.AppendChild(rowItem);
                    }

                    var rowItemTotal = new RowItem { ItemType = ItemValues.Grand };
                    rowItemTotal.AppendChild(new MemberPropertyIndex());
                    columnItems.AppendChild(rowItemTotal);
                }
            }

            if (pt.Values.Count() > 1)
            {
                // -2 is the sentinal value for "Values"
                if (pt.ColumnLabels.Any(cl => cl.SourceName == XLConstants.PivotTableValuesSentinalLabel))
                    columnFields.AppendChild(new Field { Index = -2 });
                else if (pt.RowLabels.Any(rl => rl.SourceName == XLConstants.PivotTableValuesSentinalLabel))
                {
                    pivotTableDefinition.DataOnRows = true;
                    rowFields.AppendChild(new Field { Index = -2 });
                }
            }

            foreach (var xlpf in pt.Fields)
            {
                IXLPivotField labelField = null;
                var pf = new PivotField { ShowAll = false, Name = xlpf.CustomName };

                if (pt.RowLabels.Any(p => p.SourceName == xlpf.SourceName))
                {
                    labelField = pt.RowLabels.Single(p => p.SourceName == xlpf.SourceName);
                    pf.Axis = PivotTableAxisValues.AxisRow;
                }
                else if (pt.ColumnLabels.Any(p => p.SourceName == xlpf.SourceName))
                {
                    labelField = pt.ColumnLabels.Single(p => p.SourceName == xlpf.SourceName);
                    pf.Axis = PivotTableAxisValues.AxisColumn;
                }
                else if (pt.ReportFilters.Any(p => p.SourceName == xlpf.SourceName))
                {
                    location.ColumnsPerPage = 1;
                    location.RowPageCount = 1;
                    pf.Axis = PivotTableAxisValues.AxisPage;
                    pageFields.AppendChild(new PageField { Hierarchy = -1, Field = pt.Fields.IndexOf(xlpf) });
                }

                if (pt.Values.Any(p => p.SourceName == xlpf.SourceName))
                    pf.DataField = true;

                var fieldItems = new Items();

                if (xlpf.SharedStrings.Count > 0)
                {
                    for (uint i = 0; i < xlpf.SharedStrings.Count; i++)
                    {
                        var item = new Item { Index = i };
                        if (labelField != null && labelField.Collapsed)
                            item.HideDetails = BooleanValue.FromBoolean(false);
                        fieldItems.AppendChild(item);
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

                fieldItems.Count = Convert.ToUInt32(fieldItems.Count());
                pf.AppendChild(fieldItems);
                pivotFields.AppendChild(pf);
            }

            pivotTableDefinition.AppendChild(location);
            pivotTableDefinition.AppendChild(pivotFields);

            if (pt.RowLabels.Any())
            {
                rowFields.Count = Convert.ToUInt32(rowFields.Count());
                pivotTableDefinition.AppendChild(rowFields);
            }
            else
            {
                rowItems.AppendChild(new RowItem());
            }

            rowItems.Count = Convert.ToUInt32(rowItems.Count());
            pivotTableDefinition.AppendChild(rowItems);

            if (!pt.ColumnLabels.Any(cl => cl.CustomName != XLConstants.PivotTableValuesSentinalLabel))
            {
                for (int i = 0; i < pt.Values.Count(); i++)
                {
                    var rowItem = new RowItem();
                    rowItem.Index = Convert.ToUInt32(i);
                    rowItem.AppendChild(new MemberPropertyIndex() { Val = i });
                    columnItems.AppendChild(rowItem);
                }
            }

            if (columnFields.Any())
            {
                columnFields.Count = Convert.ToUInt32(columnFields.Count());
                pivotTableDefinition.AppendChild(columnFields);
            }

            if (columnItems.Any())
            {
                columnItems.Count = Convert.ToUInt32(columnItems.Count());
                pivotTableDefinition.AppendChild(columnItems);
            }

            if (pt.ReportFilters.Any())
            {
                pageFields.Count = Convert.ToUInt32(pageFields.Count());
                pivotTableDefinition.AppendChild(pageFields);
            }


            var dataFields = new DataFields();
            foreach (var value in pt.Values)
            {
                var sourceColumn =
                    pt.SourceRange.Columns().FirstOrDefault(c => c.Cell(1).Value.ToString() == value.SourceName);
                if (sourceColumn == null) continue;

                UInt32 numberFormatId = 0;
                if (value.NumberFormat.NumberFormatId != -1 || context.SharedNumberFormats.ContainsKey(value.NumberFormat.NumberFormatId))
                    numberFormatId = (UInt32)value.NumberFormat.NumberFormatId;
                else if (context.SharedNumberFormats.Any(snf => snf.Value.NumberFormat.Format == value.NumberFormat.Format))
                    numberFormatId = (UInt32)context.SharedNumberFormats.First(snf => snf.Value.NumberFormat.Format == value.NumberFormat.Format).Key;

                var df = new DataField
                {
                    Name = value.CustomName,
                    Field = (UInt32)sourceColumn.ColumnNumber() - 1,
                    Subtotal = value.SummaryFormula.ToOpenXml(),
                    ShowDataAs = value.Calculation.ToOpenXml(),
                    NumberFormatId = numberFormatId
                };

                if (!String.IsNullOrEmpty(value.BaseField))
                {
                    var baseField = pt.SourceRange.Columns().FirstOrDefault(c => c.Cell(1).Value.ToString() == value.BaseField);
                    if (baseField != null)
                    {
                        df.BaseField = baseField.ColumnNumber() - 1;

                        var items = baseField.CellsUsed()
                            .Select(c => c.Value)
                            .Skip(1) // Skip header column
                            .Distinct().ToList();

                        if (items.Any(i => i.Equals(value.BaseItem)))
                            df.BaseItem = Convert.ToUInt32(items.IndexOf(value.BaseItem));
                    }
                }
                else
                {
                    df.BaseField = 0;
                }

                if (value.CalculationItem == XLPivotCalculationItem.Previous)
                    df.BaseItem = 1048828U;
                else if (value.CalculationItem == XLPivotCalculationItem.Next)
                    df.BaseItem = 1048829U;
                else if (df.BaseItem == null || !df.BaseItem.HasValue)
                    df.BaseItem = 0U;

                dataFields.AppendChild(df);
            }

            dataFields.Count = Convert.ToUInt32(dataFields.Count());
            pivotTableDefinition.AppendChild(dataFields);

            pivotTableDefinition.AppendChild(new PivotTableStyle
            {
                Name = Enum.GetName(typeof(XLPivotTableTheme), pt.Theme),
                ShowRowHeaders = pt.ShowRowHeaders,
                ShowColumnHeaders = pt.ShowColumnHeaders,
                ShowRowStripes = pt.ShowRowStripes,
                ShowColumnStripes = pt.ShowColumnStripes
            });

            #region Excel 2010 Features

            var pivotTableDefinitionExtensionList = new PivotTableDefinitionExtensionList();

            var pivotTableDefinitionExtension = new PivotTableDefinitionExtension
            { Uri = "{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}" };
            pivotTableDefinitionExtension.AddNamespaceDeclaration("x14",
                "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");

            var pivotTableDefinition2 = new DocumentFormat.OpenXml.Office2010.Excel.PivotTableDefinition
            { EnableEdit = pt.EnableCellEditing, HideValuesRow = !pt.ShowValuesRow };
            pivotTableDefinition2.AddNamespaceDeclaration("xm", "http://schemas.microsoft.com/office/excel/2006/main");

            pivotTableDefinitionExtension.AppendChild(pivotTableDefinition2);

            pivotTableDefinitionExtensionList.AppendChild(pivotTableDefinitionExtension);
            pivotTableDefinition.AppendChild(pivotTableDefinitionExtensionList);

            #endregion

            pivotTablePart.PivotTableDefinition = pivotTableDefinition;
        }


        private static void GenerateWorksheetCommentsPartContent(WorksheetCommentsPart worksheetCommentsPart,
            XLWorksheet xlWorksheet)
        {
            var comments = new Comments();
            var commentList = new CommentList();
            var authorsDict = new Dictionary<String, Int32>();
            foreach (var c in xlWorksheet.Internals.CellsCollection.GetCells(c => c.HasComment))
            {
                var comment = new Comment { Reference = c.Address.ToStringRelative() };
                var authorName = c.Comment.Author;

                Int32 authorId;
                if (!authorsDict.TryGetValue(authorName, out authorId))
                {
                    authorId = authorsDict.Count;
                    authorsDict.Add(authorName, authorId);
                }
                comment.AuthorId = (UInt32)authorId;

                var commentText = new CommentText();
                foreach (var rt in c.Comment)
                {
                    commentText.Append(GetRun(rt));
                }

                comment.Append(commentText);
                commentList.Append(comment);
            }

            var authors = new Authors();
            foreach (var author in authorsDict.Select(a => new Author { Text = a.Key }))
            {
                authors.Append(author);
            }
            comments.Append(authors);
            comments.Append(commentList);

            worksheetCommentsPart.Comments = comments;
        }

        // Generates content of vmlDrawingPart1.
        private static void GenerateVmlDrawingPartContent(VmlDrawingPart vmlDrawingPart, XLWorksheet xlWorksheet,
            SaveContext context)
        {
            var ms = new MemoryStream();
            CopyStream(vmlDrawingPart.GetStream(FileMode.OpenOrCreate), ms);
            ms.Position = 0;
            var writer = new XmlTextWriter(vmlDrawingPart.GetStream(FileMode.Create), Encoding.UTF8);

            writer.WriteStartElement("xml");

            const string shapeTypeId = "_x0000_t202"; // arbitrary, assigned by office

            new Vml.Shapetype(
                new Vml.Stroke { JoinStyle = Vml.StrokeJoinStyleValues.Miter },
                new Vml.Path { AllowGradientShape = true, ConnectionPointType = ConnectValues.Rectangle }
                )
            {
                Id = shapeTypeId,
                CoordinateSize = "21600,21600",
                OptionalNumber = 202,
                EdgePath = "m,l,21600r21600,l21600,xe",
            }
                .WriteTo(writer);

            var cellWithComments = xlWorksheet.Internals.CellsCollection.GetCells().Where(c => c.HasComment);

            foreach (var c in cellWithComments)
            {
                GenerateShape(c, shapeTypeId).WriteTo(writer);
            }

            if (ms.Length > 0)
            {
                ms.Position = 0;
                var xdoc = XDocumentExtensions.Load(ms);
                xdoc.Root.Elements().ForEach(e => writer.WriteRaw(e.ToString()));
            }


            writer.WriteEndElement();
            writer.Flush();
            writer.Close();
        }

        // VML Shape for Comment
        private static Vml.Shape GenerateShape(XLCell c, string shapeTypeId)
        {
            var rowNumber = c.Address.RowNumber;
            var columnNumber = c.Address.ColumnNumber;

            var shapeId = String.Format("_x0000_s{0}", c.Comment.ShapeId);
            // Unique per cell (workbook?), e.g.: "_x0000_s1026"
            var anchor = GetAnchor(c);
            var textBox = GetTextBox(c.Comment.Style);
            var fill = new Vml.Fill { Color2 = "#" + c.Comment.Style.ColorsAndLines.FillColor.Color.ToHex().Substring(2) };
            if (c.Comment.Style.ColorsAndLines.FillTransparency < 1)
                fill.Opacity =
                    Math.Round(Convert.ToDouble(c.Comment.Style.ColorsAndLines.FillTransparency), 2).ToString(
                        CultureInfo.InvariantCulture);
            var stroke = GetStroke(c);
            var shape = new Vml.Shape(
                fill,
                stroke,
                new Vml.Shadow { On = true, Color = "black", Obscured = true },
                new Vml.Path { ConnectionPointType = ConnectValues.None },
                textBox,
                new ClientData(
                    new MoveWithCells(c.Comment.Style.Properties.Positioning == XLDrawingAnchor.Absolute
                        ? "True"
                        : "False"), // Counterintuitive
                    new ResizeWithCells(c.Comment.Style.Properties.Positioning == XLDrawingAnchor.MoveAndSizeWithCells
                        ? "False"
                        : "True"), // Counterintuitive
                    anchor,
                    new HorizontalTextAlignment(c.Comment.Style.Alignment.Horizontal.ToString().ToCamel()),
                    new Vml.Spreadsheet.VerticalTextAlignment(c.Comment.Style.Alignment.Vertical.ToString().ToCamel()),
                    new AutoFill("False"),
                    new CommentRowTarget { Text = (rowNumber - 1).ToString() },
                    new CommentColumnTarget { Text = (columnNumber - 1).ToString() },
                    new Locked(c.Comment.Style.Protection.Locked ? "True" : "False"),
                    new LockText(c.Comment.Style.Protection.LockText ? "True" : "False"),
                    new Visible(c.Comment.Visible ? "True" : "False")
                    )
                { ObjectType = ObjectValues.Note }
                )
            {
                Id = shapeId,
                Type = "#" + shapeTypeId,
                Style = GetCommentStyle(c),
                FillColor = "#" + c.Comment.Style.ColorsAndLines.FillColor.Color.ToHex().Substring(2),
                StrokeColor = "#" + c.Comment.Style.ColorsAndLines.LineColor.Color.ToHex().Substring(2),
                StrokeWeight = String.Format(CultureInfo.InvariantCulture, "{0}pt", c.Comment.Style.ColorsAndLines.LineWeight),
                InsetMode = c.Comment.Style.Margins.Automatic ? InsetMarginValues.Auto : InsetMarginValues.Custom
            };
            if (!XLHelper.IsNullOrWhiteSpace(c.Comment.Style.Web.AlternateText))
                shape.Alternate = c.Comment.Style.Web.AlternateText;


            return shape;
        }

        private static Vml.Stroke GetStroke(XLCell c)
        {
            var lineDash = c.Comment.Style.ColorsAndLines.LineDash;
            var stroke = new Vml.Stroke
            {
                LineStyle = c.Comment.Style.ColorsAndLines.LineStyle.ToOpenXml(),
                DashStyle =
                    lineDash == XLDashStyle.RoundDot || lineDash == XLDashStyle.SquareDot
                        ? "shortDot"
                        : lineDash.ToString().ToCamel()
            };
            if (lineDash == XLDashStyle.RoundDot)
                stroke.EndCap = Vml.StrokeEndCapValues.Round;
            if (c.Comment.Style.ColorsAndLines.LineTransparency < 1)
                stroke.Opacity =
                    Math.Round(Convert.ToDouble(c.Comment.Style.ColorsAndLines.LineTransparency), 2).ToString(
                        CultureInfo.InvariantCulture);
            return stroke;
        }

        private static void AddPictureAnchor(WorksheetPart worksheetPart, Drawings.IXLPicture picture)
        {
            var drawingsPart = worksheetPart.DrawingsPart ??
                               worksheetPart.AddNewPart<DrawingsPart>(
                                    GeneratePartId(picture.Name, worksheetPart));

            if (drawingsPart.WorksheetDrawing == null)
            {
                drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
            }

            var worksheetDrawing = drawingsPart.WorksheetDrawing;

            var imagePart = drawingsPart.AddImagePart(picture.GetImagePartType(),
                                                        GeneratePartId(picture.Name, drawingsPart));

            using (Stream stream = new MemoryStream())
            {
                picture.ImageStream.CopyTo(stream);
                stream.Seek(0, SeekOrigin.Begin);
                imagePart.FeedData(stream);
            }

            var extentsCx = picture.Width;
            var extentsCy = picture.Height;

            var nvps = worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>();
            var nvpId = nvps.Count() > 0 ?
                (UInt32Value)worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>().Max(p => p.Id.Value) + 1 :
                1U;
            if (picture.IsAbsolute)
            {
                Xdr.AbsoluteAnchor absoluteAnchor;
                absoluteAnchor = new Xdr.AbsoluteAnchor(
                    new Xdr.Position
                    {
                        X = picture.OffsetX,
                        Y = picture.OffsetY
                    },
                    new Xdr.Extent
                    {
                        Cx = extentsCx,
                        Cy = extentsCy
                    },
                    new Xdr.Picture(
                        new Xdr.NonVisualPictureProperties(
                            new Xdr.NonVisualDrawingProperties { Id = nvpId, Name = picture.Name },
                            new Xdr.NonVisualPictureDrawingProperties(new PictureLocks { NoChangeAspect = true, NoMove = true, NoResize = true })
                        ),
                        new Xdr.BlipFill(
                            new Blip { Embed = drawingsPart.GetIdOfPart(imagePart), CompressionState = BlipCompressionValues.Print },
                            new Stretch(new FillRectangle())
                        ),
                        new Xdr.ShapeProperties(
                            new Transform2D(
                                new Offset { X = 0, Y = 0 },
                                new Extents { Cx = extentsCx, Cy = extentsCy }
                            ),
                            new PresetGeometry { Preset = ShapeTypeValues.Rectangle }
                        )
                    ),
                    new Xdr.ClientData()
                );

                worksheetDrawing.Append(absoluteAnchor);
            }
            else
            {
                var markers = picture.GetMarkers();
                Xdr.FromMarker fMark;
                Xdr.ToMarker tMark;
                if (markers.Count == 2)
                {
                    fMark = new Xdr.FromMarker
                    {
                        ColumnId = new Xdr.ColumnId(markers[0].GetZeroBasedColumn().ToString()),
                        RowId = new Xdr.RowId(markers[0].GetZeroBasedRow().ToString()),
                        ColumnOffset = new Xdr.ColumnOffset((markers[0].ColumnOffset + picture.OffsetX).ToString()),
                        RowOffset = new Xdr.RowOffset((markers[0].RowOffset + picture.OffsetY).ToString())
                    };
                    tMark = new Xdr.ToMarker
                    {
                        ColumnId = new Xdr.ColumnId(markers[1].GetZeroBasedColumn().ToString()),
                        RowId = new Xdr.RowId(markers[1].GetZeroBasedRow().ToString()),
                        ColumnOffset = new Xdr.ColumnOffset((markers[1].ColumnOffset + picture.OffsetX).ToString()),
                        RowOffset = new Xdr.RowOffset((markers[1].RowOffset + picture.OffsetY).ToString())
                    };

                    Xdr.TwoCellAnchor cellAnchor;
                    cellAnchor = new Xdr.TwoCellAnchor(
                        fMark,
                        tMark,
                        new Xdr.Picture(
                            new Xdr.NonVisualPictureProperties(
                                new Xdr.NonVisualDrawingProperties { Id = nvpId, Name = picture.Name },
                                new Xdr.NonVisualPictureDrawingProperties(new PictureLocks { NoChangeAspect = true, NoMove = true, NoResize = true })
                            ),
                            new Xdr.BlipFill(
                                new Blip { Embed = drawingsPart.GetIdOfPart(imagePart), CompressionState = BlipCompressionValues.Print },
                                new Stretch(new FillRectangle())
                            ),
                            new Xdr.ShapeProperties(
                                new Transform2D(
                                    new Offset { X = 0, Y = 0 },
                                    new Extents { Cx = extentsCx, Cy = extentsCy }
                                ),
                                new PresetGeometry { Preset = ShapeTypeValues.Rectangle }
                            )
                        ),
                        new Xdr.ClientData()
                    );

                    worksheetDrawing.Append(cellAnchor);
                }
                else if (markers.Count == 1)
                {
                    fMark = new Xdr.FromMarker
                    {
                        ColumnId = new Xdr.ColumnId(markers[0].GetZeroBasedColumn().ToString()),
                        RowId = new Xdr.RowId(markers[0].GetZeroBasedRow().ToString()),
                        ColumnOffset = new Xdr.ColumnOffset((markers[0].ColumnOffset + picture.OffsetX).ToString()),
                        RowOffset = new Xdr.RowOffset((markers[0].RowOffset + picture.OffsetY).ToString())
                    };

                    Xdr.OneCellAnchor cellAnchor;
                    cellAnchor = new Xdr.OneCellAnchor(
                        fMark,
                        new Xdr.Extent
                        {
                            Cx = extentsCx,
                            Cy = extentsCy
                        },
                        new Xdr.Picture(
                            new Xdr.NonVisualPictureProperties(
                                new Xdr.NonVisualDrawingProperties { Id = nvpId, Name = picture.Name },
                                new Xdr.NonVisualPictureDrawingProperties(new PictureLocks { NoChangeAspect = true, NoMove = true, NoResize = true })
                            ),
                            new Xdr.BlipFill(
                                new Blip { Embed = drawingsPart.GetIdOfPart(imagePart), CompressionState = BlipCompressionValues.Print },
                                new Stretch(new FillRectangle())
                            ),
                            new Xdr.ShapeProperties(
                                new Transform2D(
                                    new Offset { X = 0, Y = 0 },
                                    new Extents { Cx = extentsCx, Cy = extentsCy }
                                ),
                                new PresetGeometry { Preset = ShapeTypeValues.Rectangle }
                            )
                        ),
                        new Xdr.ClientData()
                    );

                    worksheetDrawing.Append(cellAnchor);
                }
            }
        }

        private static Regex embedRegex = new Regex("[^a-zA-Z0-9]");

        public static string GeneratePartId(string name, OpenXmlPart oxp)
        {
            var partId = name ?? "rId1";
            partId = embedRegex.Replace(partId, "");

            // We guarantee the id uniqueness based off the name
            try
            {
                oxp.GetPartById(partId);
            }
            catch(ArgumentOutOfRangeException)
            {
                return partId;
            }

            partId += "c";
            return GeneratePartId(partId, oxp);
        }

        private static Vml.TextBox GetTextBox(IXLDrawingStyle ds)
        {
            var sb = new StringBuilder();
            var a = ds.Alignment;

            if (a.Direction == XLDrawingTextDirection.Context)
                sb.Append("mso-direction-alt:auto;");
            else if (a.Direction == XLDrawingTextDirection.RightToLeft)
                sb.Append("direction:RTL;");

            if (a.Orientation != XLDrawingTextOrientation.LeftToRight)
            {
                sb.Append("layout-flow:vertical;");
                if (a.Orientation == XLDrawingTextOrientation.BottomToTop)
                    sb.Append("mso-layout-flow-alt:bottom-to-top;");
                else if (a.Orientation == XLDrawingTextOrientation.Vertical)
                    sb.Append("mso-layout-flow-alt:top-to-bottom;");
            }
            if (a.AutomaticSize)
                sb.Append("mso-fit-shape-to-text:t;");
            var retVal = new Vml.TextBox { Style = sb.ToString() };
            var dm = ds.Margins;
            if (!dm.Automatic)
                retVal.Inset = String.Format("{0}in,{1}in,{2}in,{3}in",
                    dm.Left.ToInvariantString(),
                    dm.Top.ToInvariantString(),
                    dm.Right.ToInvariantString(),
                    dm.Bottom.ToInvariantString());

            return retVal;
        }

        private static Anchor GetAnchor(XLCell cell)
        {
            var c = cell.Comment;
            var cWidth = c.Style.Size.Width;
            var fcNumber = c.Position.Column - 1;
            var fcOffset = Convert.ToInt32(c.Position.ColumnOffset * 7.5);
            var widthFromColumns = cell.Worksheet.Column(c.Position.Column).Width - c.Position.ColumnOffset;
            var lastCell = cell.CellRight(c.Position.Column - cell.Address.ColumnNumber);
            while (widthFromColumns <= cWidth)
            {
                lastCell = lastCell.CellRight();
                widthFromColumns += lastCell.WorksheetColumn().Width;
            }

            var lcNumber = lastCell.WorksheetColumn().ColumnNumber() - 1;
            var lcOffset = Convert.ToInt32((lastCell.WorksheetColumn().Width - (widthFromColumns - cWidth)) * 7.5);

            var cHeight = c.Style.Size.Height; //c.Style.Size.Height * 72.0;
            var frNumber = c.Position.Row - 1;
            var frOffset = Convert.ToInt32(c.Position.RowOffset);
            var heightFromRows = cell.Worksheet.Row(c.Position.Row).Height - c.Position.RowOffset;
            lastCell = cell.CellBelow(c.Position.Row - cell.Address.RowNumber);
            while (heightFromRows <= cHeight)
            {
                lastCell = lastCell.CellBelow();
                heightFromRows += lastCell.WorksheetRow().Height;
            }

            var lrNumber = lastCell.WorksheetRow().RowNumber() - 1;
            var lrOffset = Convert.ToInt32(lastCell.WorksheetRow().Height - (heightFromRows - cHeight));
            return new Anchor
            {
                Text = string.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}",
                    fcNumber, fcOffset,
                    frNumber, frOffset,
                    lcNumber, lcOffset,
                    lrNumber, lrOffset
                    )
            };
        }

        private static StringValue GetCommentStyle(XLCell cell)
        {
            var c = cell.Comment;
            var sb = new StringBuilder("position:absolute; ");

            sb.Append("visibility:");
            sb.Append(c.Visible ? "visible" : "hidden");
            sb.Append(";");

            sb.Append("width:");
            sb.Append(Math.Round(c.Style.Size.Width * 7.5, 2).ToInvariantString());
            sb.Append("pt;");
            sb.Append("height:");
            sb.Append(Math.Round(c.Style.Size.Height, 2).ToInvariantString());
            sb.Append("pt;");

            sb.Append("z-index:");
            sb.Append(c.ZOrder.ToString());


            return sb.ToString();
        }

        #region GenerateWorkbookStylesPartContent

        private void GenerateWorkbookStylesPartContent(WorkbookStylesPart workbookStylesPart, SaveContext context)
        {
            var defaultStyle = new XLStyle(null, DefaultStyle);
            var defaultStyleId = GetStyleId(defaultStyle);
            if (!context.SharedFonts.ContainsKey(defaultStyle.Font))
                context.SharedFonts.Add(defaultStyle.Font, new FontInfo { FontId = 0, Font = defaultStyle.Font as XLFont });

            var sharedFills = new Dictionary<IXLFill, FillInfo>
            {{defaultStyle.Fill, new FillInfo {FillId = 2, Fill = defaultStyle.Fill as XLFill}}};

            var sharedBorders = new Dictionary<IXLBorder, BorderInfo>
            {{defaultStyle.Border, new BorderInfo {BorderId = 0, Border = defaultStyle.Border as XLBorder}}};

            var sharedNumberFormats = new Dictionary<IXLNumberFormatBase, NumberFormatInfo>
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

            // To determine the default workbook style, we look for the style with builtInId = 0 (I hope that is the correct approach)
            UInt32 defaultFormatId;
            if (workbookStylesPart.Stylesheet.CellStyles.Elements<CellStyle>().Any(c => c.BuiltinId != null && c.BuiltinId.HasValue && c.BuiltinId.Value == 0))
            {
                // Possible to have duplicate default cell styles - occurs when file gets saved under different cultures.
                // We prefer the style that is named Normal
                var normalCellStyles = workbookStylesPart.Stylesheet.CellStyles.Elements<CellStyle>()
                    .Where(c => c.BuiltinId != null && c.BuiltinId.HasValue && c.BuiltinId.Value == 0)
                    .OrderBy(c => c.Name != null && c.Name.HasValue && c.Name.Value == "Normal");

                defaultFormatId = normalCellStyles.Last().FormatId.Value;
            }
            else if (workbookStylesPart.Stylesheet.CellStyles.Elements<CellStyle>().Any())
                defaultFormatId = workbookStylesPart.Stylesheet.CellStyles.Elements<CellStyle>().Max(c => c.FormatId.Value) + 1;
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
            var numberFormatCount = 1;
            var xlStyles = new HashSet<Int32>();
            var pivotTableNumberFormats = new HashSet<IXLPivotValueFormat>();

            foreach (var worksheet in WorksheetsInternal)
            {
                foreach (var s in worksheet.GetStyleIds().Where(s => !xlStyles.Contains(s)))
                    xlStyles.Add(s);

                foreach (
                    var s in
                        worksheet.Internals.ColumnsCollection.Select(kp => kp.Value.GetStyleId()).Where(
                            s => !xlStyles.Contains(s)))
                    xlStyles.Add(s);

                foreach (
                    var s in
                        worksheet.Internals.RowsCollection.Select(kp => kp.Value.GetStyleId()).Where(
                            s => !xlStyles.Contains(s))
                    )
                    xlStyles.Add(s);

                foreach (var ptnf in worksheet.PivotTables.SelectMany(pt => pt.Values.Select(ptv => ptv.NumberFormat)).Distinct().Where(nf => !pivotTableNumberFormats.Contains(nf)))
                    pivotTableNumberFormats.Add(ptnf);
            }

            foreach (var numberFormat in pivotTableNumberFormats)
            {
                if (numberFormat.NumberFormatId != -1
                    || sharedNumberFormats.ContainsKey(numberFormat))
                    continue;

                sharedNumberFormats.Add(numberFormat,
                    new NumberFormatInfo
                    {
                        NumberFormatId = XLConstants.NumberOfBuiltInStyles + numberFormatCount,
                        NumberFormat = numberFormat
                    });
                numberFormatCount++;
            }

            foreach (var xlStyle in xlStyles.Select(GetStyleById))
            {
                if (!context.SharedFonts.ContainsKey(xlStyle.Font))
                    context.SharedFonts.Add(xlStyle.Font,
                        new FontInfo { FontId = fontCount++, Font = xlStyle.Font as XLFont });

                if (!sharedFills.ContainsKey(xlStyle.Fill))
                    sharedFills.Add(xlStyle.Fill, new FillInfo { FillId = fillCount++, Fill = xlStyle.Fill as XLFill });

                if (!sharedBorders.ContainsKey(xlStyle.Border))
                    sharedBorders.Add(xlStyle.Border,
                        new BorderInfo { BorderId = borderCount++, Border = xlStyle.Border as XLBorder });

                if (xlStyle.NumberFormat.NumberFormatId != -1
                    || sharedNumberFormats.ContainsKey(xlStyle.NumberFormat))
                    continue;

                sharedNumberFormats.Add(xlStyle.NumberFormat,
                    new NumberFormatInfo
                    {
                        NumberFormatId = XLConstants.NumberOfBuiltInStyles + numberFormatCount,
                        NumberFormat = xlStyle.NumberFormat
                    });
                numberFormatCount++;
            }

            var allSharedNumberFormats = ResolveNumberFormats(workbookStylesPart, sharedNumberFormats, defaultFormatId);
            foreach (var nf in allSharedNumberFormats)
            {
                context.SharedNumberFormats.Add(nf.Value.NumberFormatId, nf.Value);
            }

            ResolveFonts(workbookStylesPart, context);
            var allSharedFills = ResolveFills(workbookStylesPart, sharedFills);
            var allSharedBorders = ResolveBorders(workbookStylesPart, sharedBorders);

            foreach (var id in xlStyles)
            {
                var xlStyle = GetStyleById(id);
                if (context.SharedStyles.ContainsKey(id)) continue;

                var numberFormatId = xlStyle.NumberFormat.NumberFormatId >= 0
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

            if (!workbookStylesPart.Stylesheet.CellStyles.Elements<CellStyle>().Any(c => c.BuiltinId != null && c.BuiltinId.HasValue && c.BuiltinId.Value == 0U))
                workbookStylesPart.Stylesheet.CellStyles.AppendChild(new CellStyle { Name = "Normal", FormatId = defaultFormatId, BuiltinId = 0U });

            workbookStylesPart.Stylesheet.CellStyles.Count = (UInt32)workbookStylesPart.Stylesheet.CellStyles.Count();

            var newSharedStyles = new Dictionary<Int32, StyleInfo>();
            foreach (var ss in context.SharedStyles)
            {
                var styleId = -1;
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

            AddDifferentialFormats(workbookStylesPart, context);
        }

        private void AddDifferentialFormats(WorkbookStylesPart workbookStylesPart, SaveContext context)
        {
            if (workbookStylesPart.Stylesheet.DifferentialFormats == null)
                workbookStylesPart.Stylesheet.DifferentialFormats = new DifferentialFormats();


            var differentialFormats = workbookStylesPart.Stylesheet.DifferentialFormats;

            FillDifferentialFormatsCollection(differentialFormats, context.DifferentialFormats);


            foreach (var ws in Worksheets)
            {
                foreach (var cf in ws.ConditionalFormats)
                {
                    if (!context.DifferentialFormats.ContainsKey(cf.Style))
                        AddDifferentialFormat(workbookStylesPart.Stylesheet.DifferentialFormats, cf, context);
                }
            }

            differentialFormats.Count = (UInt32)differentialFormats.Count();
            if (differentialFormats.Count == 0)
                workbookStylesPart.Stylesheet.DifferentialFormats = null;
        }

        private void FillDifferentialFormatsCollection(DifferentialFormats differentialFormats,
            Dictionary<IXLStyle, int> dictionary)
        {
            dictionary.Clear();
            var id = 0;
            foreach (var df in differentialFormats.Elements<DifferentialFormat>())
            {
                var style = new XLStyle(new XLStylizedEmpty(DefaultStyle), DefaultStyle);
                LoadFont(df.Font, style.Font);
                LoadBorder(df.Border, style.Border);
                LoadNumberFormat(df.NumberingFormat, style.NumberFormat);
                LoadFill(df.Fill, style.Fill);
                if (!dictionary.ContainsKey(style))
                    dictionary.Add(style, ++id);
            }
        }

        private static void AddDifferentialFormat(DifferentialFormats differentialFormats, IXLConditionalFormat cf,
            SaveContext context)
        {
            var differentialFormat = new DifferentialFormat();
            differentialFormat.Append(GetNewFont(new FontInfo { Font = cf.Style.Font as XLFont }, false));
            if (!XLHelper.IsNullOrWhiteSpace(cf.Style.NumberFormat.Format))
            {
                var numberFormat = new NumberingFormat
                {
                    NumberFormatId = (UInt32)(XLConstants.NumberOfBuiltInStyles + differentialFormats.Count()),
                    FormatCode = cf.Style.NumberFormat.Format
                };
                differentialFormat.Append(numberFormat);
            }
            differentialFormat.Append(GetNewFill(new FillInfo { Fill = cf.Style.Fill as XLFill }, false));
            differentialFormat.Append(GetNewBorder(new BorderInfo { Border = cf.Style.Border as XLBorder }, false));

            differentialFormats.Append(differentialFormat);

            context.DifferentialFormats.Add(cf.Style, differentialFormats.Count() - 1);
        }

        private static void ResolveRest(WorkbookStylesPart workbookStylesPart, SaveContext context)
        {
            if (workbookStylesPart.Stylesheet.CellFormats == null)
                workbookStylesPart.Stylesheet.CellFormats = new CellFormats();

            foreach (var styleInfo in context.SharedStyles.Values)
            {
                var info = styleInfo;
                var foundOne =
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

            foreach (var styleInfo in context.SharedStyles.Values)
            {
                var info = styleInfo;
                var foundOne =
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
                ApplyNumberFormat = true,
                ApplyAlignment = true,
                ApplyFill = ApplyFill(styleInfo),
                ApplyBorder = ApplyBorder(styleInfo),
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
                f.BorderId != null && styleInfo.BorderId == f.BorderId
                && f.FillId != null && styleInfo.FillId == f.FillId
                && f.FontId != null && styleInfo.FontId == f.FontId
                && f.NumberFormatId != null && styleInfo.NumberFormatId == f.NumberFormatId
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
            foreach (var borderInfo in sharedBorders.Values)
            {
                var borderId = 0;
                var foundOne = false;
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
                    new BorderInfo { Border = borderInfo.Border, BorderId = (UInt32)borderId });
            }
            workbookStylesPart.Stylesheet.Borders.Count = (UInt32)workbookStylesPart.Stylesheet.Borders.Count();
            return allSharedBorders;
        }

        private static Border GetNewBorder(BorderInfo borderInfo, Boolean ignoreMod = true)
        {
            var border = new Border();
            if (borderInfo.Border.DiagonalUpModified || ignoreMod)
                border.DiagonalUp = borderInfo.Border.DiagonalUp;

            if (borderInfo.Border.DiagonalDownModified || ignoreMod)
                border.DiagonalDown = borderInfo.Border.DiagonalDown;

            if (borderInfo.Border.LeftBorderModified || borderInfo.Border.LeftBorderColorModified || ignoreMod)
            {
                var leftBorder = new LeftBorder { Style = borderInfo.Border.LeftBorder.ToOpenXml() };
                if (borderInfo.Border.LeftBorderColorModified || ignoreMod)
                {
                    var leftBorderColor = GetNewColor(borderInfo.Border.LeftBorderColor);
                    leftBorder.AppendChild(leftBorderColor);
                }
                border.AppendChild(leftBorder);
            }

            if (borderInfo.Border.RightBorderModified || borderInfo.Border.RightBorderColorModified || ignoreMod)
            {
                var rightBorder = new RightBorder { Style = borderInfo.Border.RightBorder.ToOpenXml() };
                if (borderInfo.Border.RightBorderColorModified || ignoreMod)
                {
                    var rightBorderColor = GetNewColor(borderInfo.Border.RightBorderColor);
                    rightBorder.AppendChild(rightBorderColor);
                }
                border.AppendChild(rightBorder);
            }

            if (borderInfo.Border.TopBorderModified || borderInfo.Border.TopBorderColorModified || ignoreMod)
            {
                var topBorder = new TopBorder { Style = borderInfo.Border.TopBorder.ToOpenXml() };
                if (borderInfo.Border.TopBorderColorModified || ignoreMod)
                {
                    var topBorderColor = GetNewColor(borderInfo.Border.TopBorderColor);
                    topBorder.AppendChild(topBorderColor);
                }
                border.AppendChild(topBorder);
            }

            if (borderInfo.Border.BottomBorderModified || borderInfo.Border.BottomBorderColorModified || ignoreMod)
            {
                var bottomBorder = new BottomBorder { Style = borderInfo.Border.BottomBorder.ToOpenXml() };
                if (borderInfo.Border.BottomBorderColorModified || ignoreMod)
                {
                    var bottomBorderColor = GetNewColor(borderInfo.Border.BottomBorderColor);
                    bottomBorder.AppendChild(bottomBorderColor);
                }
                border.AppendChild(bottomBorder);
            }

            if (borderInfo.Border.DiagonalBorderModified || borderInfo.Border.DiagonalBorderColorModified || ignoreMod)
            {
                var DiagonalBorder = new DiagonalBorder { Style = borderInfo.Border.DiagonalBorder.ToOpenXml() };
                if (borderInfo.Border.DiagonalBorderColorModified || ignoreMod)
                {
                    var DiagonalBorderColor = GetNewColor(borderInfo.Border.DiagonalBorderColor);
                    DiagonalBorder.AppendChild(DiagonalBorderColor);
                }
                border.AppendChild(DiagonalBorder);
            }

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
            foreach (var fillInfo in sharedFills.Values)
            {
                var fillId = 0;
                var foundOne = false;
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
                allSharedFills.Add(fillInfo.Fill, new FillInfo { Fill = fillInfo.Fill, FillId = (UInt32)fillId });
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
            var patternFill1 = new PatternFill { PatternType = patternValues };
            fill1.AppendChild(patternFill1);
            fills.AppendChild(fill1);
        }

        private static Fill GetNewFill(FillInfo fillInfo, Boolean ignoreMod = true)
        {
            var fill = new Fill();

            var patternFill = new PatternFill();
            if (fillInfo.Fill.PatternTypeModified || ignoreMod)
                patternFill.PatternType = fillInfo.Fill.PatternType.ToOpenXml();

            if (fillInfo.Fill.PatternColorModified || ignoreMod)
            {
                var foregroundColor = new ForegroundColor();
                if (fillInfo.Fill.PatternColor.ColorType == XLColorType.Color)
                    foregroundColor.Rgb = fillInfo.Fill.PatternColor.Color.ToHex();
                else if (fillInfo.Fill.PatternColor.ColorType == XLColorType.Indexed)
                    foregroundColor.Indexed = (UInt32)fillInfo.Fill.PatternColor.Indexed;
                else
                {
                    foregroundColor.Theme = (UInt32)fillInfo.Fill.PatternColor.ThemeColor;
                    if (fillInfo.Fill.PatternColor.ThemeTint != 0)
                        foregroundColor.Tint = fillInfo.Fill.PatternColor.ThemeTint;
                }
                patternFill.AppendChild(foregroundColor);
            }

            if (fillInfo.Fill.PatternBackgroundColorModified || ignoreMod)
            {
                var backgroundColor = new BackgroundColor();
                if (fillInfo.Fill.PatternBackgroundColor.ColorType == XLColorType.Color)
                    backgroundColor.Rgb = fillInfo.Fill.PatternBackgroundColor.Color.ToHex();
                else if (fillInfo.Fill.PatternBackgroundColor.ColorType == XLColorType.Indexed)
                    backgroundColor.Indexed = (UInt32)fillInfo.Fill.PatternBackgroundColor.Indexed;
                else
                {
                    backgroundColor.Theme = (UInt32)fillInfo.Fill.PatternBackgroundColor.ThemeColor;
                    if (fillInfo.Fill.PatternBackgroundColor.ThemeTint != 0)
                        backgroundColor.Tint = fillInfo.Fill.PatternBackgroundColor.ThemeTint;
                }
                patternFill.AppendChild(backgroundColor);
            }

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
            foreach (var fontInfo in context.SharedFonts.Values)
            {
                var fontId = 0;
                var foundOne = false;
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
                newFonts.Add(fontInfo.Font, new FontInfo { Font = fontInfo.Font, FontId = (UInt32)fontId });
            }
            context.SharedFonts.Clear();
            foreach (var kp in newFonts)
                context.SharedFonts.Add(kp.Key, kp.Value);

            workbookStylesPart.Stylesheet.Fonts.Count = (UInt32)workbookStylesPart.Stylesheet.Fonts.Count();
        }

        private static Font GetNewFont(FontInfo fontInfo, Boolean ignoreMod = true)
        {
            var font = new Font();
            var bold = (fontInfo.Font.BoldModified || ignoreMod) && fontInfo.Font.Bold ? new Bold() : null;
            var italic = (fontInfo.Font.ItalicModified || ignoreMod) && fontInfo.Font.Italic ? new Italic() : null;
            var underline = (fontInfo.Font.UnderlineModified || ignoreMod) &&
                            fontInfo.Font.Underline != XLFontUnderlineValues.None
                ? new Underline { Val = fontInfo.Font.Underline.ToOpenXml() }
                : null;
            var strike = (fontInfo.Font.StrikethroughModified || ignoreMod) && fontInfo.Font.Strikethrough
                ? new Strike()
                : null;
            var verticalAlignment = fontInfo.Font.VerticalAlignmentModified || ignoreMod
                ? new VerticalTextAlignment { Val = fontInfo.Font.VerticalAlignment.ToOpenXml() }
                : null;
            var shadow = (fontInfo.Font.ShadowModified || ignoreMod) && fontInfo.Font.Shadow ? new Shadow() : null;
            var fontSize = fontInfo.Font.FontSizeModified || ignoreMod
                ? new FontSize { Val = fontInfo.Font.FontSize }
                : null;
            var color = fontInfo.Font.FontColorModified || ignoreMod ? GetNewColor(fontInfo.Font.FontColor) : null;

            var fontName = fontInfo.Font.FontNameModified || ignoreMod
                ? new FontName { Val = fontInfo.Font.FontName }
                : null;
            var fontFamilyNumbering = fontInfo.Font.FontFamilyNumberingModified || ignoreMod
                ? new FontFamilyNumbering { Val = (Int32)fontInfo.Font.FontFamilyNumbering }
                : null;

            if (bold != null)
                font.AppendChild(bold);
            if (italic != null)
                font.AppendChild(italic);
            if (underline != null)
                font.AppendChild(underline);
            if (strike != null)
                font.AppendChild(strike);
            if (verticalAlignment != null)
                font.AppendChild(verticalAlignment);
            if (shadow != null)
                font.AppendChild(shadow);
            if (fontSize != null)
                font.AppendChild(fontSize);
            if (color != null)
                font.AppendChild(color);
            if (fontName != null)
                font.AppendChild(fontName);
            if (fontFamilyNumbering != null)
                font.AppendChild(fontFamilyNumbering);

            return font;
        }

        private static Color GetNewColor(XLColor xlColor)
        {
            var color = new Color();
            if (xlColor.ColorType == XLColorType.Color)
                color.Rgb = xlColor.Color.ToHex();
            else if (xlColor.ColorType == XLColorType.Indexed)
                color.Indexed = (UInt32)xlColor.Indexed;
            else
            {
                color.Theme = (UInt32)xlColor.ThemeColor;
                if (xlColor.ThemeTint != 0)
                    color.Tint = xlColor.ThemeTint;
            }
            return color;
        }

        private static TabColor GetTabColor(XLColor xlColor)
        {
            var color = new TabColor();
            if (xlColor.ColorType == XLColorType.Color)
                color.Rgb = xlColor.Color.ToHex();
            else if (xlColor.ColorType == XLColorType.Indexed)
                color.Indexed = (UInt32)xlColor.Indexed;
            else
            {
                color.Theme = (UInt32)xlColor.ThemeColor;
                if (xlColor.ThemeTint != 0)
                    color.Tint = xlColor.ThemeTint;
            }
            return color;
        }

        private bool FontsAreEqual(Font f, IXLFont xlFont)
        {
            var nf = new XLFont { Bold = f.Bold != null, Italic = f.Italic != null };
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

        private static Dictionary<IXLNumberFormatBase, NumberFormatInfo> ResolveNumberFormats(
            WorkbookStylesPart workbookStylesPart,
            Dictionary<IXLNumberFormatBase, NumberFormatInfo> sharedNumberFormats,
            UInt32 defaultFormatId)
        {
            if (workbookStylesPart.Stylesheet.NumberingFormats == null)
            {
                workbookStylesPart.Stylesheet.NumberingFormats = new NumberingFormats();
                workbookStylesPart.Stylesheet.NumberingFormats.AppendChild(new NumberingFormat()
                {
                    NumberFormatId = 0,
                    FormatCode = ""
                });
            }

            var allSharedNumberFormats = new Dictionary<IXLNumberFormatBase, NumberFormatInfo>();
            foreach (var numberFormatInfo in sharedNumberFormats.Values.Where(nf => nf.NumberFormatId != defaultFormatId))
            {
                var numberingFormatId = XLConstants.NumberOfBuiltInStyles + 1;
                var foundOne = false;
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

        private static bool NumberFormatsAreEqual(NumberingFormat nf, IXLNumberFormatBase xlNumberFormat)
        {
            var newXLNumberFormat = new XLNumberFormat();

            if (nf.FormatCode != null && !XLHelper.IsNullOrWhiteSpace(nf.FormatCode.Value))
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

            var maxColumn = 0;

            var sheetDimensionReference = "A1";
            if (xlWorksheet.Internals.CellsCollection.Count > 0)
            {
                maxColumn = xlWorksheet.Internals.CellsCollection.MaxColumnUsed;
                var maxRow = xlWorksheet.Internals.CellsCollection.MaxRowUsed;
                sheetDimensionReference = "A1:" + XLHelper.GetColumnLetterFromNumber(maxColumn) +
                                          maxRow.ToInvariantString();
            }

            if (xlWorksheet.Internals.ColumnsCollection.Count > 0)
            {
                var maxColCollection = xlWorksheet.Internals.ColumnsCollection.Keys.Max();
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

            if (xlWorksheet.TabSelected)
                sheetView.TabSelected = true;
            else
                sheetView.TabSelected = null;

            if (xlWorksheet.RightToLeft)
                sheetView.RightToLeft = true;
            else
                sheetView.RightToLeft = null;

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

            if (xlWorksheet.SheetView.View == XLSheetViewOptions.Normal)
                sheetView.View = null;
            else
                sheetView.View = xlWorksheet.SheetView.View.ToOpenXml();

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

            pane.TopLeftCell = XLHelper.GetColumnLetterFromNumber(xlWorksheet.SheetView.SplitColumn + 1)
                               + (xlWorksheet.SheetView.SplitRow + 1);

            if (hSplit == 0 && ySplit == 0)
                sheetView.RemoveAllChildren<Pane>();

            if (xlWorksheet.SelectedRanges.Any() || xlWorksheet.ActiveCell != null)
            {
                sheetView.RemoveAllChildren<Selection>();

                var firstSelection = xlWorksheet.SelectedRanges.FirstOrDefault();
                var selection = new Selection();
                if (xlWorksheet.ActiveCell != null)
                    selection.ActiveCell = xlWorksheet.ActiveCell.Address.ToStringRelative(false);
                else if (firstSelection != null)
                    selection.ActiveCell = firstSelection.RangeAddress.FirstAddress.ToStringRelative(false);


                var seqRef = new List<String> { selection.ActiveCell.Value };
                seqRef.AddRange(xlWorksheet.SelectedRanges
                    .Select(range => range.RangeAddress.ToStringRelative(false)));


                selection.SequenceOfReferences = new ListValue<StringValue> { InnerText = String.Join(" ", seqRef.Distinct().ToArray()) };

                sheetView.Append(selection);
            }

            if (xlWorksheet.SheetView.ZoomScale == 100)
                sheetView.ZoomScale = null;
            else
                sheetView.ZoomScale = (UInt32)Math.Max(10, Math.Min(400, xlWorksheet.SheetView.ZoomScale));

            if (xlWorksheet.SheetView.ZoomScaleNormal == 100)
                sheetView.ZoomScaleNormal = null;
            else
                sheetView.ZoomScaleNormal = (UInt32)Math.Max(10, Math.Min(400, xlWorksheet.SheetView.ZoomScaleNormal));

            if (xlWorksheet.SheetView.ZoomScalePageLayoutView == 100)
                sheetView.ZoomScalePageLayoutView = null;
            else
                sheetView.ZoomScalePageLayoutView = (UInt32)Math.Max(10, Math.Min(400, xlWorksheet.SheetView.ZoomScalePageLayoutView));

            if (xlWorksheet.SheetView.ZoomScaleSheetLayoutView == 100)
                sheetView.ZoomScaleSheetLayoutView = null;
            else
                sheetView.ZoomScaleSheetLayoutView = (UInt32)Math.Max(10, Math.Min(400, xlWorksheet.SheetView.ZoomScaleSheetLayoutView));

            #endregion

            var maxOutlineColumn = 0;
            if (xlWorksheet.ColumnCount() > 0)
                maxOutlineColumn = xlWorksheet.GetMaxColumnOutline();

            var maxOutlineRow = 0;
            if (xlWorksheet.RowCount() > 0)
                maxOutlineRow = xlWorksheet.GetMaxRowOutline();

            #region SheetFormatProperties

            if (worksheetPart.Worksheet.SheetFormatProperties == null)
                worksheetPart.Worksheet.SheetFormatProperties = new SheetFormatProperties();

            cm.SetElement(XLWSContentManager.XLWSContents.SheetFormatProperties,
                worksheetPart.Worksheet.SheetFormatProperties);

            worksheetPart.Worksheet.SheetFormatProperties.DefaultRowHeight = xlWorksheet.RowHeight.SaveRound();

            if (xlWorksheet.RowHeightChanged)
                worksheetPart.Worksheet.SheetFormatProperties.CustomHeight = true;
            else
                worksheetPart.Worksheet.SheetFormatProperties.CustomHeight = null;


            var worksheetColumnWidth = GetColumnWidth(xlWorksheet.ColumnWidth).SaveRound();
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

                var worksheetStyleId = context.SharedStyles[xlWorksheet.GetStyleId()].StyleId;
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

                for (var co = minInColumnsCollection; co <= maxInColumnsCollection; co++)
                {
                    UInt32 styleId;
                    Double columnWidth;
                    var isHidden = false;
                    var collapsed = false;
                    var outlineLevel = 0;
                    if (xlWorksheet.Internals.ColumnsCollection.ContainsKey(co))
                    {
                        styleId = context.SharedStyles[xlWorksheet.Internals.ColumnsCollection[co].GetStyleId()].StyleId;
                        columnWidth = GetColumnWidth(xlWorksheet.Internals.ColumnsCollection[co].Width).SaveRound();
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

                var collection = maxInColumnsCollection;
                foreach (
                    var col in
                        columns.Elements<Column>().Where(c => c.Min > (UInt32)(collection)).OrderBy(
                            c => c.Min.Value))
                {
                    col.Style = worksheetStyleId;
                    col.Width = worksheetColumnWidth;
                    col.CustomWidth = true;

                    if ((Int32)col.Max.Value > maxInColumnsCollection)
                        maxInColumnsCollection = (Int32)col.Max.Value;
                }

                if (maxInColumnsCollection < XLHelper.MaxColumnNumber && !xlWorksheet.Style.Equals(DefaultStyle))
                {
                    var column = new Column
                    {
                        Min = (UInt32)(maxInColumnsCollection + 1),
                        Max = (UInt32)(XLHelper.MaxColumnNumber),
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

            var lastRow = 0;
            var sheetDataRows =
                sheetData.Elements<Row>().ToDictionary(r => r.RowIndex == null ? ++lastRow : (Int32)r.RowIndex.Value,
                    r => r);
            foreach (
                var r in
                    xlWorksheet.Internals.RowsCollection.Deleted.Where(r => sheetDataRows.ContainsKey(r.Key)))
            {
                sheetData.RemoveChild(sheetDataRows[r.Key]);
                sheetDataRows.Remove(r.Key);
                xlWorksheet.Internals.CellsCollection.deleted.Remove(r.Key);
            }

            var distinctRows = xlWorksheet.Internals.CellsCollection.RowsCollection.Keys.Union(xlWorksheet.Internals.RowsCollection.Keys);
            var noRows = !sheetData.Elements<Row>().Any();
            foreach (var distinctRow in distinctRows.OrderBy(r => r))
            {
                Row row;
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
                            var minRow = sheetDataRows.Where(r => r.Key > (Int32)row.RowIndex.Value).Min(r => r.Key);
                            var rowBeforeInsert = sheetDataRows[minRow];
                            sheetData.InsertBefore(row, rowBeforeInsert);
                        }
                        else
                            sheetData.AppendChild(row);
                    }
                }

                if (maxColumn > 0)
                    row.Spans = new ListValue<StringValue> { InnerText = "1:" + maxColumn.ToInvariantString() };

                row.Height = null;
                row.CustomHeight = null;
                row.Hidden = null;
                row.StyleIndex = null;
                row.CustomFormat = null;
                row.Collapsed = null;
                if (xlWorksheet.Internals.RowsCollection.ContainsKey(distinctRow))
                {
                    var thisRow = xlWorksheet.Internals.RowsCollection[distinctRow];
                    if (thisRow.HeightChanged)
                    {
                        row.Height = thisRow.Height.SaveRound();
                        row.CustomHeight = true;
                        row.CustomFormat = true;
                    }

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

                var lastCell = 0;
                var cellsByReference = row.Elements<Cell>().ToDictionary(c => c.CellReference == null
                    ? XLHelper.GetColumnLetterFromNumber(
                        ++lastCell) + distinctRow
                    : c.CellReference.Value, c => c);

                foreach (var kpDel in xlWorksheet.Internals.CellsCollection.deleted.ToList())
                {
                    foreach (var delCo in kpDel.Value.ToList())
                    {
                        var key = XLHelper.GetColumnLetterFromNumber(delCo) + kpDel.Key.ToInvariantString();
                        if (!cellsByReference.ContainsKey(key)) continue;
                        row.RemoveChild(cellsByReference[key]);
                        kpDel.Value.Remove(delCo);
                    }
                    if (kpDel.Value.Count == 0)
                        xlWorksheet.Internals.CellsCollection.deleted.Remove(kpDel.Key);
                }


                if (!xlWorksheet.Internals.CellsCollection.RowsCollection.ContainsKey(distinctRow)) continue;

                var isNewRow = !row.Elements<Cell>().Any();
                lastCell = 0;
                var mRows = row.Elements<Cell>().ToDictionary(c => XLHelper.GetColumnNumberFromAddress(c.CellReference == null
                    ? (XLHelper.GetColumnLetterFromNumber(++lastCell) + distinctRow) : c.CellReference.Value), c => c);
                foreach (var opCell in xlWorksheet.Internals.CellsCollection.RowsCollection[distinctRow].Values
                    .OrderBy(c => c.Address.ColumnNumber)
                    .Select(c => c))
                {
                    var styleId = context.SharedStyles[opCell.GetStyleId()].StyleId;

                    var dataType = opCell.DataType;
                    var cellReference = (opCell.Address).GetTrimmedAddress();

                    var isEmpty = opCell.IsEmpty(true);

                    Cell cell = null;
                    if (cellsByReference.ContainsKey(cellReference))
                    {
                        cell = cellsByReference[cellReference];
                        if (isEmpty)
                        {
                            cell.Remove();
                        }
                    }

                    if (!isEmpty)
                    {
                        if (cell == null)
                        {
                            cell = new Cell();
                            cell.CellReference = new StringValue(cellReference);

                            if (isNewRow)
                                row.AppendChild(cell);
                            else
                            {
                                var newColumn = XLHelper.GetColumnNumberFromAddress(cellReference);

                                Cell cellBeforeInsert = null;
                                int[] lastCo = { Int32.MaxValue };
                                foreach (var c in mRows.Where(kp => kp.Key > newColumn).Where(c => lastCo[0] > c.Key))
                                {
                                    cellBeforeInsert = c.Value;
                                    lastCo[0] = c.Key;
                                }
                                if (cellBeforeInsert == null)
                                    row.AppendChild(cell);
                                else
                                    row.InsertBefore(cell, cellBeforeInsert);
                            }
                        }

                        cell.StyleIndex = styleId;
                        var formula = opCell.FormulaA1;
                        if (opCell.HasFormula)
                        {
                            if (formula.StartsWith("{"))
                            {
                                formula = formula.Substring(1, formula.Length - 2);
                                var f = new CellFormula { FormulaType = CellFormulaValues.Array };

                                if (opCell.FormulaReference == null)
                                    opCell.FormulaReference = opCell.AsRange().RangeAddress;
                                if (opCell.FormulaReference.FirstAddress.Equals(opCell.Address))
                                {
                                    f.Text = formula;
                                    f.Reference = opCell.FormulaReference.ToStringRelative();
                                }

                                cell.CellFormula = f;
                            }
                            else
                            {
                                cell.CellFormula = new CellFormula();
                                cell.CellFormula.Text = formula;
                            }

                            cell.CellValue = null;
                        }
                        else
                        {
                            cell.CellFormula = null;

                            cell.DataType = opCell.DataType == XLCellValues.DateTime ? null : GetCellValue(opCell);

                            if (dataType == XLCellValues.Text)
                            {
                                if (opCell.InnerText.Length == 0)
                                    cell.CellValue = null;
                                else
                                {
                                    if (opCell.ShareString)
                                    {
                                        var cellValue = new CellValue();
                                        cellValue.Text = opCell.SharedStringId.ToString();
                                        cell.CellValue = cellValue;
                                    }
                                    else
                                    {
                                        var text = opCell.GetString();
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
                                var cellValue = new CellValue();
                                cellValue.Text =
                                    XLCell.BaseDate.Add(timeSpan).ToOADate().ToInvariantString();
                                cell.CellValue = cellValue;
                            }
                            else if (dataType == XLCellValues.DateTime || dataType == XLCellValues.Number)
                            {
                                if (!XLHelper.IsNullOrWhiteSpace(opCell.InnerText))
                                {
                                    var cellValue = new CellValue();
                                    cellValue.Text = Double.Parse(opCell.InnerText, XLHelper.NumberStyle, XLHelper.ParseCulture).ToInvariantString();
                                    cell.CellValue = cellValue;
                                }
                            }
                            else
                            {
                                var cellValue = new CellValue();
                                cellValue.Text = opCell.InnerText;
                                cell.CellValue = cellValue;
                            }
                        }
                    }
                }
                xlWorksheet.Internals.CellsCollection.deleted.Remove(distinctRow);
            }
            foreach (
                var r in
                    xlWorksheet.Internals.CellsCollection.deleted.Keys.Where(
                        sheetDataRows.ContainsKey))
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
                if (!XLHelper.IsNullOrWhiteSpace(protection.PasswordHash))
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

            worksheetPart.Worksheet.RemoveAllChildren<AutoFilter>();
            if (xlWorksheet.AutoFilter.Enabled)
            {
                var previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.AutoFilter);
                worksheetPart.Worksheet.InsertAfter(new AutoFilter(), previousElement);


                var autoFilter = worksheetPart.Worksheet.Elements<AutoFilter>().First();
                cm.SetElement(XLWSContentManager.XLWSContents.AutoFilter, autoFilter);

                PopulateAutoFilter(xlWorksheet.AutoFilter, autoFilter);
            }
            else
            {
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

                foreach (var mergeCell in (xlWorksheet).Internals.MergedRanges.Select(
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

            #region Conditional Formatting

            if (!xlWorksheet.ConditionalFormats.Any())
            {
                worksheetPart.Worksheet.RemoveAllChildren<ConditionalFormatting>();
                cm.SetElement(XLWSContentManager.XLWSContents.ConditionalFormatting, null);
            }
            else
            {
                worksheetPart.Worksheet.RemoveAllChildren<ConditionalFormatting>();
                var previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.ConditionalFormatting);

                var priority = 1; // priority is 1 origin in Microsoft Excel
                foreach (var cfGroup in xlWorksheet.ConditionalFormats
                    .GroupBy(
                        c => c.Range.RangeAddress.ToStringRelative(false),
                        c => c,
                        (key, g) => new { RangeId = key, CfList = g.ToList() }
                    )
                    )
                {
                    var conditionalFormatting = new ConditionalFormatting
                    {
                        SequenceOfReferences =
                            new ListValue<StringValue> { InnerText = cfGroup.RangeId }
                    };
                    foreach (var cf in cfGroup.CfList)
                    {
                        conditionalFormatting.Append(XLCFConverters.Convert(cf, priority, context));
                        priority++;
                    }
                    worksheetPart.Worksheet.InsertAfter(conditionalFormatting, previousElement);
                    previousElement = conditionalFormatting;
                    cm.SetElement(XLWSContentManager.XLWSContents.ConditionalFormatting, conditionalFormatting);
                }
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
                foreach (var dv in xlWorksheet.DataValidations)
                {
                    var sequence = dv.Ranges.Aggregate(String.Empty, (current, r) => current + (r.RangeAddress + " "));

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
                foreach (var hl in xlWorksheet.Hyperlinks)
                {
                    Hyperlink hyperlink;
                    if (hl.IsExternal)
                    {
                        var rId = context.RelIdGenerator.GetNext(RelType.Workbook);
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
                    if (!XLHelper.IsNullOrWhiteSpace(hl.Tooltip))
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

                if (xlWorksheet.PageSetup.PagesWide >= 0 && xlWorksheet.PageSetup.PagesWide != 1)
                    pageSetup.FitToWidth = (UInt32)xlWorksheet.PageSetup.PagesWide;

                if (xlWorksheet.PageSetup.PagesTall >= 0 && xlWorksheet.PageSetup.PagesTall != 1)
                    pageSetup.FitToHeight = (UInt32)xlWorksheet.PageSetup.PagesTall;
            }

            #endregion

            #region HeaderFooter

            var headerFooter = worksheetPart.Worksheet.Elements<HeaderFooter>().FirstOrDefault();
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
                headerFooter.DifferentFirst = xlWorksheet.PageSetup.DifferentFirstPageOnHF;
                headerFooter.DifferentOddEven = xlWorksheet.PageSetup.DifferentOddEvenPagesOnHF;

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

            var rowBreakCount = xlWorksheet.PageSetup.RowBreaks.Count;
            if (rowBreakCount > 0)
            {
                rowBreaks.Count = (UInt32)rowBreakCount;
                rowBreaks.ManualBreakCount = (UInt32)rowBreakCount;
                var lastRowNum = (UInt32)xlWorksheet.RangeAddress.LastAddress.RowNumber;
                foreach (var break1 in xlWorksheet.PageSetup.RowBreaks.Select(rb => new Break
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

            var columnBreakCount = xlWorksheet.PageSetup.ColumnBreaks.Count;
            if (columnBreakCount > 0)
            {
                columnBreaks.Count = (UInt32)columnBreakCount;
                columnBreaks.ManualBreakCount = (UInt32)columnBreakCount;
                var maxColumnNumber = (UInt32)xlWorksheet.RangeAddress.LastAddress.ColumnNumber;
                foreach (var break1 in xlWorksheet.PageSetup.ColumnBreaks.Select(cb => new Break
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
                var tablePart in
                    from XLTable xlTable in xlWorksheet.Tables select new TablePart { Id = xlTable.RelId })
                tableParts.AppendChild(tablePart);

            #endregion

            #region Drawings

            var pics = xlWorksheet.Pictures();
            if (pics != null)
            {
                foreach (Drawings.IXLPicture pic in pics)
                {
                    AddPictureAnchor(worksheetPart, pic);
                }
            }

            if (xlWorksheet.Pictures() != null && xlWorksheet.Pictures().Count > 0)
            {
                Drawing worksheetDrawing = new Drawing { Id = worksheetPart.GetIdOfPart(worksheetPart.DrawingsPart) };
                worksheetDrawing.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                worksheetPart.Worksheet.InsertBefore<Drawing>(worksheetDrawing, tableParts);
            }

            #endregion

            #region LegacyDrawing

            if (xlWorksheet.LegacyDrawingIsNew)
            {
                worksheetPart.Worksheet.RemoveAllChildren<LegacyDrawing>();
                {
                    if (!XLHelper.IsNullOrWhiteSpace(xlWorksheet.LegacyDrawingId))
                    {
                        var previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.LegacyDrawing);
                        worksheetPart.Worksheet.InsertAfter(new LegacyDrawing { Id = xlWorksheet.LegacyDrawingId },
                            previousElement);
                    }
                }
            }

            #endregion

            #region LegacyDrawingHeaderFooter

            //LegacyDrawingHeaderFooter legacyHeaderFooter = worksheetPart.Worksheet.Elements<LegacyDrawingHeaderFooter>().FirstOrDefault();
            //if (legacyHeaderFooter != null)
            //{
            //    worksheetPart.Worksheet.RemoveAllChildren<LegacyDrawingHeaderFooter>();
            //    {
            //            var previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.LegacyDrawingHeaderFooter);
            //            worksheetPart.Worksheet.InsertAfter(new LegacyDrawingHeaderFooter { Id = xlWorksheet.LegacyDrawingId },
            //                                                previousElement);
            //    }
            //}

            #endregion
        }

        private static void PopulateAutoFilter(XLAutoFilter xlAutoFilter, AutoFilter autoFilter)
        {
            var filterRange = xlAutoFilter.Range;
            autoFilter.Reference = filterRange.RangeAddress.ToString();

            foreach (var kp in xlAutoFilter.Filters)
            {
                var filterColumn = new FilterColumn { ColumnId = (UInt32)kp.Key - 1 };
                var xlFilterColumn = xlAutoFilter.Column(kp.Key);
                var filterType = xlFilterColumn.FilterType;
                if (filterType == XLFilterType.Custom)
                {
                    var customFilters = new CustomFilters();
                    foreach (var filter in kp.Value)
                    {
                        var customFilter = new CustomFilter { Val = filter.Value.ToString() };

                        if (filter.Operator != XLFilterOperator.Equal)
                            customFilter.Operator = filter.Operator.ToOpenXml();

                        if (filter.Connector == XLConnector.And)
                            customFilters.And = true;

                        customFilters.Append(customFilter);
                    }
                    filterColumn.Append(customFilters);
                }
                else if (filterType == XLFilterType.TopBottom)
                {
                    var top101 = new Top10 { Val = (double)xlFilterColumn.TopBottomValue };
                    if (xlFilterColumn.TopBottomType == XLTopBottomType.Percent)
                        top101.Percent = true;
                    if (xlFilterColumn.TopBottomPart == XLTopBottomPart.Bottom)
                        top101.Top = false;

                    filterColumn.Append(top101);
                }
                else if (filterType == XLFilterType.Dynamic)
                {
                    var dynamicFilter = new DynamicFilter
                    { Type = xlFilterColumn.DynamicType.ToOpenXml(), Val = xlFilterColumn.DynamicValue };
                    filterColumn.Append(dynamicFilter);
                }
                else
                {
                    var filters = new Filters();
                    foreach (var filter in kp.Value)
                    {
                        filters.Append(new Filter { Val = filter.Value.ToString() });
                    }

                    filterColumn.Append(filters);
                }
                autoFilter.Append(filterColumn);
            }


            if (xlAutoFilter.Sorted)
            {
                var sortState = new SortState
                {
                    Reference =
                        filterRange.Range(filterRange.FirstCell().CellBelow(), filterRange.LastCell()).RangeAddress.
                            ToString()
                };
                var sortCondition = new SortCondition
                {
                    Reference =
                        filterRange.Range(1, xlAutoFilter.SortColumn, filterRange.RowCount(),
                            xlAutoFilter.SortColumn).RangeAddress.ToString()
                };
                if (xlAutoFilter.SortOrder == XLSortOrder.Descending)
                    sortCondition.Descending = true;

                sortState.Append(sortCondition);
                autoFilter.Append(sortState);
            }
        }

        private static BooleanValue GetBooleanValue(bool value, bool defaultValue)
        {
            return value == defaultValue ? null : new BooleanValue(value);
        }

        private static void CollapseColumns(Columns columns, Dictionary<uint, Column> sheetColumns)
        {
            UInt32 lastMin = 1;
            var count = sheetColumns.Count;
            var arr = sheetColumns.OrderBy(kp => kp.Key).ToArray();
            // sheetColumns[kp.Key + 1]
            //Int32 i = 0;
            //foreach (KeyValuePair<uint, Column> kp in arr
            //    //.Where(kp => !(kp.Key < count && ColumnsAreEqual(kp.Value, )))
            //    )
            for (var i = 0; i < count; i++)
            {
                var kp = arr[i];
                if (i + 1 != count && ColumnsAreEqual(kp.Value, arr[i + 1].Value)) continue;

                var newColumn = (Column)kp.Value.CloneNode(true);
                newColumn.Min = lastMin;
                var newColumnMax = newColumn.Max.Value;
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
            return Math.Min(255.0, Math.Max(0.0, columnWidth + ColumnWidthOffset));
        }

        private static void UpdateColumn(Column column, Columns columns, Dictionary<uint, Column> sheetColumnsByMin)
        {
            var co = column.Min.Value;
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
                newColumn.Width = column.Width.SaveRound();
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
    }
}
