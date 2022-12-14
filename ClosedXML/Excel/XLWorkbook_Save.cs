using ClosedXML.Extensions;
using ClosedXML.Utils;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Vml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using Anchor = DocumentFormat.OpenXml.Vml.Spreadsheet.Anchor;
using BackgroundColor = DocumentFormat.OpenXml.Spreadsheet.BackgroundColor;
using Bold = DocumentFormat.OpenXml.Spreadsheet.Bold;
using Border = DocumentFormat.OpenXml.Spreadsheet.Border;
using BottomBorder = DocumentFormat.OpenXml.Spreadsheet.BottomBorder;
using Color = DocumentFormat.OpenXml.Spreadsheet.Color;
using Comment = DocumentFormat.OpenXml.Spreadsheet.Comment;
using Comments = DocumentFormat.OpenXml.Spreadsheet.Comments;
using Drawing = DocumentFormat.OpenXml.Spreadsheet.Drawing;
using Field = DocumentFormat.OpenXml.Spreadsheet.Field;
using Fill = DocumentFormat.OpenXml.Spreadsheet.Fill;
using Font = DocumentFormat.OpenXml.Spreadsheet.Font;
using FontCharSet = DocumentFormat.OpenXml.Spreadsheet.FontCharSet;
using FontFamily = DocumentFormat.OpenXml.Spreadsheet.FontFamily;
using Fonts = DocumentFormat.OpenXml.Spreadsheet.Fonts;
using FontScheme = DocumentFormat.OpenXml.Drawing.FontScheme;
using FontSize = DocumentFormat.OpenXml.Spreadsheet.FontSize;
using ForegroundColor = DocumentFormat.OpenXml.Spreadsheet.ForegroundColor;
using Format = DocumentFormat.OpenXml.Spreadsheet.Format;
using GradientFill = DocumentFormat.OpenXml.Drawing.GradientFill;
using GradientStop = DocumentFormat.OpenXml.Drawing.GradientStop;
using Italic = DocumentFormat.OpenXml.Spreadsheet.Italic;
using LeftBorder = DocumentFormat.OpenXml.Spreadsheet.LeftBorder;
using Locked = DocumentFormat.OpenXml.Vml.Spreadsheet.Locked;
using NumberingFormat = DocumentFormat.OpenXml.Spreadsheet.NumberingFormat;
using OfficeExcel = DocumentFormat.OpenXml.Office.Excel;
using Outline = DocumentFormat.OpenXml.Drawing.Outline;
using Path = System.IO.Path;
using PatternFill = DocumentFormat.OpenXml.Spreadsheet.PatternFill;
using Properties = DocumentFormat.OpenXml.ExtendedProperties.Properties;
using RightBorder = DocumentFormat.OpenXml.Spreadsheet.RightBorder;
using Run = DocumentFormat.OpenXml.Spreadsheet.Run;
using RunProperties = DocumentFormat.OpenXml.Spreadsheet.RunProperties;
using Shadow = DocumentFormat.OpenXml.Spreadsheet.Shadow;
using Strike = DocumentFormat.OpenXml.Spreadsheet.Strike;
using Table = DocumentFormat.OpenXml.Spreadsheet.Table;
using Text = DocumentFormat.OpenXml.Spreadsheet.Text;
using TopBorder = DocumentFormat.OpenXml.Spreadsheet.TopBorder;
using Underline = DocumentFormat.OpenXml.Spreadsheet.Underline;
using VerticalTextAlignment = DocumentFormat.OpenXml.Spreadsheet.VerticalTextAlignment;
using Vml = DocumentFormat.OpenXml.Vml;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using ClosedXML.Excel.Cells;
using ClosedXML.Excel.IO;

namespace ClosedXML.Excel
{
    public partial class XLWorkbook
    {
        private Boolean Validate(SpreadsheetDocument package)
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
                var message = string.Join("\r\n", errors.Select(e => string.Format("Part {0}, Path {1}: {2}", e.Part.Uri, e.Path.XPath, e.Description)).ToArray());
                throw new ApplicationException(message);
            }
            return true;
        }

        private void CreatePackage(String filePath, SpreadsheetDocumentType spreadsheetDocumentType, SaveOptions options)
        {
            var directoryName = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrWhiteSpace(directoryName)) Directory.CreateDirectory(directoryName);

            var package = File.Exists(filePath)
                ? SpreadsheetDocument.Open(filePath, true)
                : SpreadsheetDocument.Create(filePath, spreadsheetDocumentType);

            using (package)
            {
                if (package.DocumentType != spreadsheetDocumentType)
                {
                    package.ChangeDocumentType(spreadsheetDocumentType);
                }

                CreateParts(package, options);
                if (options.ValidatePackage) Validate(package);
            }
        }

        private void CreatePackage(Stream stream, bool newStream, SpreadsheetDocumentType spreadsheetDocumentType, SaveOptions options)
        {
            var package = newStream
                ? SpreadsheetDocument.Create(stream, spreadsheetDocumentType)
                : SpreadsheetDocument.Open(stream, true);

            using (package)
            {
                CreateParts(package, options);
                if (options.ValidatePackage) Validate(package);
            }
        }

        // http://blogs.msdn.com/b/vsod/archive/2010/02/05/how-to-delete-a-worksheet-from-excel-using-open-xml-sdk-2-0.aspx
        private void DeleteSheetAndDependencies(WorkbookPart wbPart, string sheetId)
        {
            //Get the SheetToDelete from workbook.xml
            Sheet worksheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Id == sheetId);
            if (worksheet == null)
                return;

            string sheetName = worksheet.Name;
            // Get the pivot Table Parts
            var pvtTableCacheParts = wbPart.PivotTableCacheDefinitionParts;
            var pvtTableCacheDefinitionPart = new Dictionary<PivotTableCacheDefinitionPart, string>();
            foreach (PivotTableCacheDefinitionPart Item in pvtTableCacheParts)
            {
                PivotCacheDefinition pvtCacheDef = Item.PivotCacheDefinition;
                //Check if this CacheSource is linked to SheetToDelete
                if (pvtCacheDef.Descendants<CacheSource>().Any(cacheSource => cacheSource.WorksheetSource?.Sheet == sheetName))
                {
                    pvtTableCacheDefinitionPart.Add(Item, Item.ToString());
                }
            }
            foreach (var Item in pvtTableCacheDefinitionPart)
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

                foreach (var Item in definedNames.OfType<DefinedName>())
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
                    calcsToDelete.Add(Item);

                foreach (CalculationCell Item in calcsToDelete)
                    Item.Remove();

                if (!calChainPart.CalculationChain.Any())
                    wbPart.DeletePart(calChainPart);
            }
        }

        // Adds child parts and generates content of the specified part.
        private void CreateParts(SpreadsheetDocument document, SaveOptions options)
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
            context.RelIdGenerator.AddExistingValues(workbookPart, this);

            var extendedFilePropertiesPart = document.ExtendedFilePropertiesPart ??
                                             document.AddNewPart<ExtendedFilePropertiesPart>(
                                                 context.RelIdGenerator.GetNext(RelType.Workbook));

            GenerateExtendedFilePropertiesPartContent(extendedFilePropertiesPart);

            GenerateWorkbookPartContent(workbookPart, options, context);

            var sharedStringTablePart = workbookPart.SharedStringTablePart ??
                                        workbookPart.AddNewPart<SharedStringTablePart>(
                                            context.RelIdGenerator.GetNext(RelType.Workbook));

            GenerateSharedStringTablePartContent(sharedStringTablePart, context);

            var workbookStylesPart = workbookPart.WorkbookStylesPart ??
                                     workbookPart.AddNewPart<WorkbookStylesPart>(
                                         context.RelIdGenerator.GetNext(RelType.Workbook));

            GenerateWorkbookStylesPartContent(workbookStylesPart, context);

            var cacheRelIds = WorksheetsInternal
                  .Cast<XLWorksheet>()
                  .SelectMany(s => s.PivotTables.Cast<XLPivotTable>().Select(pt => pt.WorkbookCacheRelId))
                  .Where(relId => !string.IsNullOrWhiteSpace(relId))
                  .Distinct();

            foreach (var relId in cacheRelIds)
            {
                if (workbookPart.GetPartById(relId) is PivotTableCacheDefinitionPart pivotTableCacheDefinitionPart)
                    pivotTableCacheDefinitionPart.PivotCacheDefinition.CacheFields.RemoveAllChildren();
            }

            foreach (var worksheet in WorksheetsInternal.Cast<XLWorksheet>().OrderBy(w => w.Position))
            {
                WorksheetPart worksheetPart;
                var wsRelId = worksheet.RelId;
                if (workbookPart.Parts.Any(p => p.RelationshipId == wsRelId))
                    worksheetPart = (WorksheetPart)workbookPart.GetPartById(wsRelId);
                else
                    worksheetPart = workbookPart.AddNewPart<WorksheetPart>(wsRelId);

                var worksheetHasComments = worksheet.Internals.CellsCollection.GetCells(c => c.HasComment).Any();

                var commentsPart = worksheetPart.WorksheetCommentsPart;
                var vmlDrawingPart = worksheetPart.VmlDrawingParts.FirstOrDefault();
                var hasAnyVmlElements = DeleteExistingComments(worksheetPart, worksheet, commentsPart, vmlDrawingPart);

                if (worksheetHasComments)
                {
                    if (commentsPart == null)
                    {
                        commentsPart = worksheetPart.AddNewPart<WorksheetCommentsPart>(context.RelIdGenerator.GetNext(RelType.Workbook));
                        commentsPart.Comments = new Comments();
                    }

                    if (vmlDrawingPart == null)
                    {
                        if (String.IsNullOrWhiteSpace(worksheet.LegacyDrawingId))
                        {
                            worksheet.LegacyDrawingId = context.RelIdGenerator.GetNext(RelType.Workbook);
                        }

                        vmlDrawingPart = worksheetPart.AddNewPart<VmlDrawingPart>(worksheet.LegacyDrawingId);
                    }

                    GenerateWorksheetCommentsPartContent(commentsPart, worksheet);
                    hasAnyVmlElements = GenerateVmlDrawingPartContent(vmlDrawingPart, worksheet);
                }
                else
                {
                    worksheet.LegacyDrawingId = null;
                }

                // Remove empty parts
                if (commentsPart != null && (commentsPart.RootElement?.ChildElements?.Count ?? 0) == 0)
                    worksheetPart.DeletePart(commentsPart);

                if (!hasAnyVmlElements && vmlDrawingPart != null)
                    worksheetPart.DeletePart(vmlDrawingPart);

                var xlTables = (XLTables)worksheet.Tables;

                // The way forward is to have 2-phase save, this is a start of that
                // concept for tables:
                //
                // Phase 1 - synchronize part existence with tables xlWorksheet, so each
                // table has a corresponding part and part that don't are deleted.
                // This phase doesn't modify the content, it only ensures that RelIds are set
                // corresponding parts exist and the parts that don't exist are removed
                SynchronizeTableParts(xlTables, worksheetPart, context);

                // Phase 2 - At this point, all pieces must have corresponding parts
                // The only way to link between parts is through RelIds that were already
                // set in phase 1. The phase 2 is all about content of individual parts.
                // Each part should have individual writer.
                GenerateTableParts(xlTables, worksheetPart, context);

                WorksheetPartWriter.GenerateWorksheetPartContent(worksheetPart, worksheet, options, context);

                if (worksheet.PivotTables.Any())
                {
                    GeneratePivotTables(workbookPart, worksheetPart, worksheet, context);
                }

                // Remove any orphaned references - maybe more types?
                foreach (var orphan in worksheetPart.Worksheet.OfType<LegacyDrawing>().Where(lg => worksheetPart.Parts.All(p => p.RelationshipId != lg.Id)))
                    worksheetPart.Worksheet.RemoveChild(orphan);

                foreach (var orphan in worksheetPart.Worksheet.OfType<Drawing>().Where(d => worksheetPart.Parts.All(p => p.RelationshipId != d.Id)))
                    worksheetPart.Worksheet.RemoveChild(orphan);
            }

            // Remove empty pivot cache part
            if (workbookPart.Workbook.PivotCaches != null && !workbookPart.Workbook.PivotCaches.Any())
                workbookPart.Workbook.RemoveChild(workbookPart.Workbook.PivotCaches);

            if (options.GenerateCalculationChain)
                GenerateCalculationChainPartContent(workbookPart, context);
            else
                DeleteCalculationChainPartContent(workbookPart, context);

            if (workbookPart.ThemePart == null)
            {
                var themePart = workbookPart.AddNewPart<ThemePart>(context.RelIdGenerator.GetNext(RelType.Workbook));
                GenerateThemePartContent(themePart);
            }

            // Custom properties
            if (CustomProperties.Any())
            {
                var customFilePropertiesPart =
                    document.CustomFilePropertiesPart ?? document.AddNewPart<CustomFilePropertiesPart>(context.RelIdGenerator.GetNext(RelType.Workbook));

                GenerateCustomFilePropertiesPartContent(customFilePropertiesPart);
            }
            else
            {
                if (document.CustomFilePropertiesPart != null)
                    document.DeletePart(document.CustomFilePropertiesPart);
            }
            SetPackageProperties(document);

            // Clear list of deleted worksheets to prevent errors on multiple saves
            worksheets.Deleted.Clear();
        }

        private bool DeleteExistingComments(WorksheetPart worksheetPart, XLWorksheet worksheet, WorksheetCommentsPart commentsPart, VmlDrawingPart vmlDrawingPart)
        {
            // Nuke existing comments
            if (commentsPart != null)
            {
                commentsPart.Comments = new Comments();
            }

            if (vmlDrawingPart == null)
                return false;

            // Nuke the VmlDrawingPart elements for comments.
            using (var vmlStream = vmlDrawingPart.GetStream(FileMode.Open))
            {
                var xdoc = XDocumentExtensions.Load(vmlStream);
                if (xdoc == null)
                    return false;

                // Remove existing shapes for comments
                xdoc.Root
                    .Elements()
                    .Where(e => e.Name.LocalName == "shapetype" && e.Attribute("id").Value == XLConstants.Comment.ShapeTypeId)
                    .Remove();

                xdoc.Root
                    .Elements()
                    .Where(e => e.Name.LocalName == "shape" && e.Attribute("type").Value == "#" + XLConstants.Comment.ShapeTypeId)
                    .Remove();

                vmlStream.Position = 0;

                using (var writer = new XmlTextWriter(vmlStream, Encoding.UTF8))
                {
                    var contents = xdoc.ToString();
                    writer.WriteRaw(contents);
                    vmlStream.SetLength(contents.Length);
                }

                return xdoc.Root.HasElements;
            }
        }

        private static void SynchronizeTableParts(XLTables tables, WorksheetPart worksheetPart, SaveContext context)
        {
            // Remove table definition parts that are not a part of workbook
            foreach (var tableDefinitionPart in worksheetPart.GetPartsOfType<TableDefinitionPart>().ToList())
            {
                var partId = worksheetPart.GetIdOfPart(tableDefinitionPart);
                var xlWorkbookContainsTable = tables.Cast<XLTable>().Any(t => t.RelId == partId);
                if (!xlWorkbookContainsTable)
                {
                    worksheetPart.DeletePart(tableDefinitionPart);
                }
            }

            foreach (var xlTable in tables.Cast<XLTable>())
            {
                if (String.IsNullOrEmpty(xlTable.RelId))
                {
                    xlTable.RelId = context.RelIdGenerator.GetNext(RelType.Workbook);
                    worksheetPart.AddNewPart<TableDefinitionPart>(xlTable.RelId);
                }
            }
        }

        private static void GenerateTableParts(XLTables tables, WorksheetPart worksheetPart, SaveContext context)
        {
            foreach (var xlTable in tables.Cast<XLTable>())
            {
                var relId = xlTable.RelId;
                var tableDefinitionPart = (TableDefinitionPart)worksheetPart.GetPartById(relId);
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

            InsertOnVtVector(vTVectorOne, "Worksheets", 0, modifiedWorksheetsCount.ToInvariantString());
            InsertOnVtVector(vTVectorOne, "Named Ranges", 2, modifiedNamedRangesCount.ToInvariantString());

            vTVectorTwo.Size = (UInt32)(modifiedNamedRangesCount + modifiedWorksheetsCount);

            foreach (
                var vTlpstr3 in modifiedWorksheets.OrderBy(w => w.Order).Select(w => new VTLPSTR { Text = w.Name }))
                vTVectorTwo.AppendChild(vTlpstr3);

            foreach (var vTlpstr7 in modifiedNamedRanges.Select(nr => new VTLPSTR { Text = nr }))
                vTVectorTwo.AppendChild(vTlpstr7);

            if (Properties.Manager != null)
            {
                if (!String.IsNullOrWhiteSpace(Properties.Manager))
                {
                    if (properties.Manager == null)
                        properties.Manager = new Manager();

                    properties.Manager.Text = Properties.Manager;
                }
                else
                    properties.Manager = null;
            }

            if (Properties.Company == null) return;

            if (!String.IsNullOrWhiteSpace(Properties.Company))
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

        private void GenerateWorkbookPartContent(WorkbookPart workbookPart, SaveOptions options, SaveContext context)
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

            workbook.WorkbookProperties.Date1904 = OpenXmlHelper.GetBooleanValue(this.Use1904DateSystem, false);

            if (options.FilterPrivacy.HasValue)
                workbook.WorkbookProperties.FilterPrivacy = OpenXmlHelper.GetBooleanValue(options.FilterPrivacy.Value, false);

            #endregion WorkbookProperties

            #region FileSharing

            if (workbook.FileSharing == null)
                workbook.FileSharing = new FileSharing();

            workbook.FileSharing.ReadOnlyRecommended = OpenXmlHelper.GetBooleanValue(this.FileSharing.ReadOnlyRecommended, false);
            workbook.FileSharing.UserName = String.IsNullOrWhiteSpace(this.FileSharing.UserName) ? null : StringValue.FromString(this.FileSharing.UserName);

            if (!workbook.FileSharing.HasChildren && !workbook.FileSharing.HasAttributes)
                workbook.FileSharing = null;

            #endregion FileSharing

            #region WorkbookProtection

            if (this.Protection.IsProtected)
            {
                if (workbook.WorkbookProtection == null)
                    workbook.WorkbookProtection = new WorkbookProtection();

                var workbookProtection = workbook.WorkbookProtection;

                var protection = this.Protection;

                workbookProtection.WorkbookPassword = null;
                workbookProtection.WorkbookAlgorithmName = null;
                workbookProtection.WorkbookHashValue = null;
                workbookProtection.WorkbookSpinCount = null;
                workbookProtection.WorkbookSaltValue = null;

                if (protection.Algorithm == XLProtectionAlgorithm.Algorithm.SimpleHash)
                {
                    if (!String.IsNullOrWhiteSpace(protection.PasswordHash))
                        workbookProtection.WorkbookPassword = protection.PasswordHash;
                }
                else
                {
                    workbookProtection.WorkbookAlgorithmName = DescribedEnumParser<XLProtectionAlgorithm.Algorithm>.ToDescription(protection.Algorithm);
                    workbookProtection.WorkbookHashValue = protection.PasswordHash;
                    workbookProtection.WorkbookSpinCount = protection.SpinCount;
                    workbookProtection.WorkbookSaltValue = protection.Base64EncodedSalt;
                }

                workbookProtection.LockStructure = OpenXmlHelper.GetBooleanValue(!protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Structure), false);
                workbookProtection.LockWindows = OpenXmlHelper.GetBooleanValue(!protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Windows), false);
            }
            else
            {
                workbook.WorkbookProtection = null;
            }

            #endregion WorkbookProtection

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
                if (xlSheet.SheetId == 0 && String.IsNullOrWhiteSpace(xlSheet.RelId))
                {
                    rId = context.RelIdGenerator.GetNext(RelType.Workbook);

                    while (WorksheetsInternal.Cast<XLWorksheet>().Any(w => w.SheetId == Int32.Parse(rId.Substring(3))))
                        rId = context.RelIdGenerator.GetNext(RelType.Workbook);

                    xlSheet.SheetId = Int32.Parse(rId.Substring(3));
                    xlSheet.RelId = rId;
                }
                else
                {
                    if (String.IsNullOrWhiteSpace(xlSheet.RelId))
                    {
                        rId = String.Concat("rId", xlSheet.SheetId);
                        context.RelIdGenerator.AddValues(new List<String> { rId }, RelType.Workbook);
                    }
                    else
                        rId = xlSheet.RelId;
                }

                if (workbook.Sheets.Cast<Sheet>().All(s => s.Id != rId))
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
                    else
                        sheet.State = null;

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
                UInt32? firstActiveTab = null;
                UInt32? firstSelectedTab = null;
                foreach (var ws in worksheets)
                {
                    if (ws.TabActive)
                    {
                        firstActiveTab = (UInt32)(ws.Position - 1);
                        break;
                    }

                    if (ws.TabSelected)
                    {
                        firstSelectedTab = (UInt32)(ws.Position - 1);
                    }
                }

                activeTab = firstActiveTab
                         ?? firstSelectedTab
                         ?? firstSheetVisible;
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
                            (worksheetName.EscapeSheetName() + "!" +
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

                if (worksheet.AutoFilter.IsEnabled)
                {
                    var definedName = new DefinedName
                    {
                        Name = "_xlnm._FilterDatabase",
                        LocalSheetId = sheetId,
                        Text = worksheet.Name.EscapeSheetName() + "!" +
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

                    if (!String.IsNullOrWhiteSpace(nr.Comment))
                        definedName.Comment = nr.Comment;
                    definedNames.AppendChild(definedName);
                }

                var definedNameTextRow = String.Empty;
                var definedNameTextColumn = String.Empty;
                if (worksheet.PageSetup.FirstRowToRepeatAtTop > 0)
                {
                    definedNameTextRow = worksheet.Name.EscapeSheetName() + "!" + worksheet.PageSetup.FirstRowToRepeatAtTop
                                         + ":" + worksheet.PageSetup.LastRowToRepeatAtTop;
                }
                if (worksheet.PageSetup.FirstColumnToRepeatAtLeft > 0)
                {
                    var minColumn = worksheet.PageSetup.FirstColumnToRepeatAtLeft;
                    var maxColumn = worksheet.PageSetup.LastColumnToRepeatAtLeft;
                    definedNameTextColumn = worksheet.Name.EscapeSheetName() + "!" +
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

            foreach (var nr in NamedRanges.OfType<XLNamedRange>())
            {
                var refersTo = string.Join(",", nr.RangeList
                    .Select(r => r.StartsWith("#REF!") ? "#REF!" : r));

                var definedName = new DefinedName
                {
                    Name = nr.Name,
                    Text = refersTo
                };

                if (!nr.Visible)
                    definedName.Hidden = BooleanValue.FromBoolean(true);

                if (!String.IsNullOrWhiteSpace(nr.Comment))
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

            static bool HasSharedString(XLCell c)
            {
                if (c.DataType == XLDataType.Text && c.ShareString)
                    return c.StyleValue.IncludeQuotePrefix || String.IsNullOrWhiteSpace(c.FormulaA1) && c.GetText().Length > 0;
                else
                    return false;
            }

            foreach (var c in Worksheets.Cast<XLWorksheet>().SelectMany(w => w.Internals.CellsCollection.GetCells(HasSharedString)))
            {
                if (c.HasRichText)
                {
                    if (newRichStrings.TryGetValue(c.GetRichText(), out int id))
                        c.SharedStringId = id;
                    else
                    {
                        var sharedStringItem = new SharedStringItem();
                        PopulatedRichTextElements(sharedStringItem, c, context);

                        sharedStringTablePart.SharedStringTable.Append(sharedStringItem);
                        sharedStringTablePart.SharedStringTable.Count += 1;
                        sharedStringTablePart.SharedStringTable.UniqueCount += 1;

                        newRichStrings.Add(c.GetRichText(), stringId);
                        c.SharedStringId = stringId;

                        stringId++;
                    }
                }
                else
                {
                    var value = c.Value.GetText();
                    if (newStrings.TryGetValue(value, out int id))
                        c.SharedStringId = id;
                    else
                    {
                        var s = value;
                        var sharedStringItem = new SharedStringItem();
                        var text = new Text { Text = XmlEncoder.EncodeString(s) };
                        if (!s.Trim().Equals(s))
                            text.Space = SpaceProcessingModeValues.Preserve;
                        sharedStringItem.Append(text);
                        sharedStringTablePart.SharedStringTable.Append(sharedStringItem);
                        sharedStringTablePart.SharedStringTable.Count += 1;
                        sharedStringTablePart.SharedStringTable.UniqueCount += 1;

                        newStrings.Add(value, stringId);
                        c.SharedStringId = stringId;

                        stringId++;
                    }
                }
            }
        }

        internal static void PopulatedRichTextElements(RstType rstType, IXLCell cell, SaveContext context)
        {
            var richText = cell.GetRichText();
            foreach (var rt in richText.Where(r => !String.IsNullOrEmpty(r.Text)))
            {
                rstType.Append(GetRun(rt));
            }

            if (richText.HasPhonetics)
            {
                foreach (var p in richText.Phonetics)
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
                    rstType.Append(phoneticRun);
                }

                var fontKey = XLFont.GenerateKey(richText.Phonetics);
                var f = XLFontValue.FromKey(ref fontKey);

                if (!context.SharedFonts.TryGetValue(f, out FontInfo fi))
                {
                    fi = new FontInfo { Font = f };
                    context.SharedFonts.Add(f, fi);
                }

                var phoneticProperties = new PhoneticProperties
                {
                    FontId = fi.FontId
                };

                if (richText.Phonetics.Alignment != XLPhoneticAlignment.Left)
                    phoneticProperties.Alignment = richText.Phonetics.Alignment.ToOpenXml();

                if (richText.Phonetics.Type != XLPhoneticType.FullWidthKatakana)
                    phoneticProperties.Type = richText.Phonetics.Type.ToOpenXml();

                rstType.Append(phoneticProperties);
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
            var color = new Color().FromClosedXMLColor<Color>(rt.FontColor);
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

        private void DeleteCalculationChainPartContent(WorkbookPart workbookPart, SaveContext context)
        {
            if (workbookPart.CalculationChainPart != null)
                workbookPart.DeletePart(workbookPart.CalculationChainPart);
        }

        private void GenerateCalculationChainPartContent(WorkbookPart workbookPart, SaveContext context)
        {
            if (workbookPart.CalculationChainPart == null)
                workbookPart.AddNewPart<CalculationChainPart>(context.RelIdGenerator.GetNext(RelType.Workbook));

            if (workbookPart.CalculationChainPart.CalculationChain == null)
                workbookPart.CalculationChainPart.CalculationChain = new CalculationChain();

            var calculationChain = workbookPart.CalculationChainPart.CalculationChain;
            calculationChain.RemoveAllChildren<CalculationCell>();

            foreach (var worksheet in WorksheetsInternal)
            {
                foreach (var c in worksheet.Internals.CellsCollection.GetCells().Where(c => c.HasFormula))
                {
                    if (c.HasArrayFormula)
                    {
                        if (c.FormulaReference == null)
                            c.FormulaReference = c.AsRange().RangeAddress;

                        if (c.FormulaReference.FirstAddress.Equals(c.Address))
                        {
                            var cc = new CalculationCell
                            {
                                CellReference = c.Address.ToString(),
                                SheetId = worksheet.SheetId
                            };

                            cc.Array = true;
                            calculationChain.AppendChild(cc);

                            foreach (var childCell in worksheet.Range(c.FormulaReference.ToString()).Cells())
                            {
                                calculationChain.AppendChild(
                                    new CalculationCell
                                    {
                                        CellReference = childCell.Address.ToString(),
                                        SheetId = worksheet.SheetId,
                                        InChildChain = true
                                    }
                                );
                            }
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

        private void GenerateCustomFilePropertiesPartContent(CustomFilePropertiesPart customFilePropertiesPart)
        {
            var properties = new DocumentFormat.OpenXml.CustomProperties.Properties();
            properties.AddNamespaceDeclaration("vt",
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
                properties.AppendChild(customDocumentProperty);
            }

            customFilePropertiesPart.Properties = properties;
        }

        private void SetPackageProperties(OpenXmlPackage document)
        {
            var created = Properties.Created == DateTime.MinValue ? DateTime.Now : Properties.Created;
            var modified = Properties.Modified == DateTime.MinValue ? DateTime.Now : Properties.Modified;
            document.PackageProperties.Created = created;
            document.PackageProperties.Modified = modified;

#if true // Workaround: https://github.com/OfficeDev/Open-XML-SDK/issues/235

            if (Properties.LastModifiedBy == null) document.PackageProperties.LastModifiedBy = "";
            if (Properties.Author == null) document.PackageProperties.Creator = "";
            if (Properties.Title == null) document.PackageProperties.Title = "";
            if (Properties.Subject == null) document.PackageProperties.Subject = "";
            if (Properties.Category == null) document.PackageProperties.Category = "";
            if (Properties.Keywords == null) document.PackageProperties.Keywords = "";
            if (Properties.Comments == null) document.PackageProperties.Description = "";
            if (Properties.Status == null) document.PackageProperties.ContentStatus = "";

#endif

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

        private static void GenerateTableDefinitionPartContent(TableDefinitionPart tableDefinitionPart, XLTable xlTable, SaveContext context)
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

            var tableColumns = new TableColumns { Count = (UInt32)xlTable.ColumnCount() };

            UInt32 columnId = 0;
            foreach (var xlField in xlTable.Fields)
            {
                columnId++;
                var fieldName = xlField.Name;
                var tableColumn = new TableColumn
                {
                    Id = columnId,
                    Name = fieldName.Replace("_x000a_", "_x005f_x000a_").Replace(Environment.NewLine, "_x000a_")
                };

                // https://github.com/ClosedXML/ClosedXML/issues/513
                if (xlField.IsConsistentStyle())
                {
                    var style = (xlField.Column.Cells()
                        .Skip(xlTable.ShowHeaderRow ? 1 : 0)
                        .First()
                        .Style as XLStyle).Value;

                    if (!DefaultStyleValue.Equals(style) && context.DifferentialFormats.TryGetValue(style, out Int32 id))
                        tableColumn.DataFormatId = UInt32Value.FromUInt32(Convert.ToUInt32(id));
                }
                else
                    tableColumn.DataFormatId = null;

                if (xlField.IsConsistentFormula())
                {
                    string formula = xlField.Column.Cells()
                        .Skip(xlTable.ShowHeaderRow ? 1 : 0)
                        .First()
                        .FormulaA1;

                    while (formula.StartsWith("=") && formula.Length > 1)
                        formula = formula.Substring(1);

                    if (!String.IsNullOrWhiteSpace(formula))
                    {
                        tableColumn.CalculatedColumnFormula = new CalculatedColumnFormula
                        {
                            Text = formula
                        };
                    }
                }
                else
                    tableColumn.CalculatedColumnFormula = null;

                if (xlTable.ShowTotalsRow)
                {
                    if (xlField.TotalsRowFunction != XLTotalsRowFunction.None)
                    {
                        tableColumn.TotalsRowFunction = xlField.TotalsRowFunction.ToOpenXml();

                        if (xlField.TotalsRowFunction == XLTotalsRowFunction.Custom)
                            tableColumn.TotalsRowFormula = new TotalsRowFormula(xlField.TotalsRowFormulaA1);
                    }

                    if (!String.IsNullOrWhiteSpace(xlField.TotalsRowLabel))
                        tableColumn.TotalsRowLabel = xlField.TotalsRowLabel;
                }
                tableColumns.AppendChild(tableColumn);
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

                WorksheetPartWriter.PopulateAutoFilter(xlTable.AutoFilter, autoFilter1);

                table.AppendChild(autoFilter1);
            }

            table.AppendChild(tableColumns);
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
                context.PivotTables.Clear();

                // TODO: Avoid duplicate pivot caches of same source range

                PivotCache pivotCache;
                PivotTableCacheDefinitionPart pivotTableCacheDefinitionPart;
                if (!String.IsNullOrWhiteSpace(pt.WorkbookCacheRelId))
                {
                    pivotCache = pivotCaches.Cast<PivotCache>().Single(pc => pc.Id.Value == pt.WorkbookCacheRelId);
                    pivotTableCacheDefinitionPart = workbookPart.GetPartById(pt.WorkbookCacheRelId) as PivotTableCacheDefinitionPart;
                }
                else
                {
                    var workbookCacheRelId = context.RelIdGenerator.GetNext(RelType.Workbook);
                    pt.WorkbookCacheRelId = workbookCacheRelId;
                    pivotCache = new PivotCache { CacheId = cacheId++, Id = workbookCacheRelId };
                    pivotCaches.AppendChild(pivotCache);
                    pivotTableCacheDefinitionPart = workbookPart.AddNewPart<PivotTableCacheDefinitionPart>(workbookCacheRelId);
                }

                GeneratePivotTableCacheDefinitionPartContent(pivotTableCacheDefinitionPart, pt, context);

                PivotTablePart pivotTablePart;
                var createNewPivotTablePart = String.IsNullOrWhiteSpace(pt.RelId);
                if (createNewPivotTablePart)
                {
                    var relId = context.RelIdGenerator.GetNext(RelType.Workbook);
                    pt.RelId = relId;
                    pivotTablePart = worksheetPart.AddNewPart<PivotTablePart>(relId);
                }
                else
                    pivotTablePart = worksheetPart.GetPartById(pt.RelId) as PivotTablePart;

                GeneratePivotTablePartContent(pivotTablePart, pt, pivotCache.CacheId, context);

                if (createNewPivotTablePart)
                    pivotTablePart.AddPart(pivotTableCacheDefinitionPart, context.RelIdGenerator.GetNext(RelType.Workbook));
            }
        }

        // Generates content of pivotTableCacheDefinitionPart
        private static void GeneratePivotTableCacheDefinitionPartContent(
            PivotTableCacheDefinitionPart pivotTableCacheDefinitionPart, XLPivotTable pt,
            SaveContext context)
        {
            var pivotCacheDefinition = pivotTableCacheDefinitionPart.PivotCacheDefinition;
            if (pivotCacheDefinition == null)
            {
                pivotCacheDefinition = new PivotCacheDefinition { Id = "rId1" };

                pivotCacheDefinition.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                pivotTableCacheDefinitionPart.PivotCacheDefinition = pivotCacheDefinition;
            }

            #region CreatedVersion

            byte createdVersion = XLConstants.PivotTable.CreatedVersion;

            if (pivotCacheDefinition.CreatedVersion?.HasValue ?? false)
                pivotCacheDefinition.CreatedVersion = Math.Max(createdVersion, pivotCacheDefinition.CreatedVersion.Value);
            else
                pivotCacheDefinition.CreatedVersion = createdVersion;

            #endregion CreatedVersion

            #region RefreshedVersion

            byte refreshedVersion = XLConstants.PivotTable.RefreshedVersion;
            if (pivotCacheDefinition.RefreshedVersion?.HasValue ?? false)
                pivotCacheDefinition.RefreshedVersion = Math.Max(refreshedVersion, pivotCacheDefinition.RefreshedVersion.Value);
            else
                pivotCacheDefinition.RefreshedVersion = refreshedVersion;

            #endregion RefreshedVersion

            #region MinRefreshableVersion

            byte minRefreshableVersion = 3;
            if (pivotCacheDefinition.MinRefreshableVersion?.HasValue ?? false)
                pivotCacheDefinition.MinRefreshableVersion = Math.Max(minRefreshableVersion, pivotCacheDefinition.MinRefreshableVersion.Value);
            else
                pivotCacheDefinition.MinRefreshableVersion = minRefreshableVersion;

            #endregion MinRefreshableVersion

            pivotCacheDefinition.SaveData = pt.SaveSourceData;
            pivotCacheDefinition.RefreshOnLoad = true; //pt.RefreshDataOnOpen

            var pti = new PivotTableInfo
            {
                Guid = pt.Guid,
                Fields = new Dictionary<String, PivotTableFieldInfo>()
            };

            var source = pt.SourceRange;
            if (pt.ItemsToRetainPerField == XLItemsToRetain.None)
                pivotCacheDefinition.MissingItemsLimit = 0U;
            else if (pt.ItemsToRetainPerField == XLItemsToRetain.Max)
                pivotCacheDefinition.MissingItemsLimit = XLHelper.MaxRowNumber;

            // Begin CacheSource
            var cacheSource = new CacheSource { Type = SourceValues.Worksheet };
            var worksheetSource = new WorksheetSource();

            switch (pt.SourceType)
            {
                case XLPivotTableSourceType.Range:
                    worksheetSource.Name = null;
                    worksheetSource.Reference = source.RangeAddress.ToStringRelative(includeSheet: false);

                    // Do not quote worksheet name with whitespace here - issue #955
                    worksheetSource.Sheet = source.RangeAddress.Worksheet.Name;
                    break;

                case XLPivotTableSourceType.Table:
                    worksheetSource.Name = pt.SourceTable.Name;
                    worksheetSource.Reference = null;
                    worksheetSource.Sheet = null;
                    break;

                default:
                    throw new NotSupportedException($"Pivot table source type {pt.SourceType} is not supported.");
            }

            cacheSource.AppendChild(worksheetSource);
            pivotCacheDefinition.CacheSource = cacheSource;

            // End CacheSource

            // Begin CacheFields
            var cacheFields = pivotCacheDefinition.CacheFields;
            if (cacheFields == null)
            {
                cacheFields = new CacheFields();
                pivotCacheDefinition.CacheFields = cacheFields;
            }

            foreach (var c in source.Columns())
            {
                var columnNumber = c.ColumnNumber();
                var columnName = c.FirstCell().Value.ToString(CultureInfo.CurrentCulture);

                CacheField cacheField = null;

                // .CacheFields is cleared when workbook is begin saved
                // So if there are any entries, it would be from previous pivot tables
                // with an identical source range.
                // When pivot sources get its refactoring, this will not be necessary
                if (cacheFields != null)
                {
                    cacheField = pivotCacheDefinition
                        .CacheFields
                        .Elements<CacheField>()
                        .FirstOrDefault(f => f.Name == columnName);
                }

                if (cacheField == null)
                {
                    cacheField = new CacheField
                    {
                        Name = columnName,
                        SharedItems = new SharedItems()
                    };
                    cacheFields.AppendChild(cacheField);
                }
                var sharedItems = cacheField.SharedItems;

                XLPivotField xlpf;
                if (pt.Fields.Contains(columnName))
                    xlpf = pt.Fields.Get(columnName) as XLPivotField;
                else
                    xlpf = pt.Fields.Add(columnName) as XLPivotField;

                var field = pt.RowLabels
                    .Union(pt.ColumnLabels)
                    .Union(pt.ReportFilters)
                    .FirstOrDefault(f => f.SourceName == columnName);

                if (field == null)
                {
                    xlpf.ShowBlankItems = true;
                }
                else
                {
                    xlpf.CustomName = field.CustomName;
                    xlpf.SortType = field.SortType;
                    xlpf.SubtotalCaption = field.SubtotalCaption;
                    xlpf.IncludeNewItemsInFilter = field.IncludeNewItemsInFilter;
                    xlpf.Outline = field.Outline;
                    xlpf.Compact = field.Compact;
                    xlpf.SubtotalsAtTop = field.SubtotalsAtTop;
                    xlpf.RepeatItemLabels = field.RepeatItemLabels;
                    xlpf.InsertBlankLines = field.InsertBlankLines;
                    xlpf.ShowBlankItems = field.ShowBlankItems;
                    xlpf.InsertPageBreaks = field.InsertPageBreaks;
                    xlpf.Collapsed = field.Collapsed;
                    xlpf.Subtotals.AddRange(field.Subtotals);
                }

                var ptfi = new PivotTableFieldInfo
                {
                    IsTotallyBlankField = false
                };

                var sourceHeaderRow = source.FirstRow().RowNumber();
                var cellsUsed = source.CellsUsed(XLCellsUsedOptions.All,
                                                 cell => cell.Address.ColumnNumber == columnNumber
                                                         && cell.Address.RowNumber > sourceHeaderRow)
                                                 .ToArray();
                var fieldValueCells = cellsUsed.Where(cell => !cell.IsEmpty()).ToArray();
                var types = fieldValueCells.Select(cell => cell.DataType).Distinct().ToArray();
                var containsBlank = cellsUsed.Any(cell => cell.IsEmpty());

                // For a totally blank column, we need to check that all cells in column are unused
                if (fieldValueCells.Length == 0)
                {
                    ptfi.IsTotallyBlankField = true;
                    containsBlank = true;
                }

                if (types.Length > 0)
                {
                    ptfi.MixedDataType = types.Length > 1;
                    ptfi.DistinctValues = fieldValueCells
                        .Select(cell => cell.CachedValue)
                        .Distinct(XLCellValueComparer.OrdinalIgnoreCase)
                        .ToArray();

                    pti.Fields.Add(xlpf.SourceName, ptfi);

                    // If this cache field exists and contains shared items,
                    // then we can assume that this as been populated by a previous pivot table
                    if (sharedItems.Any())
                        continue;

                    // Else we have to populate the items
                    if (types.Length == 1 && types.Single() == XLDataType.Number)
                    {
                        sharedItems.ContainsSemiMixedTypes = containsBlank;
                        sharedItems.ContainsString = false;
                        sharedItems.ContainsNumber = true;

                        var allInteger = ptfi.DistinctValues.All(v => v.GetNumber() % 1 == 0);
                        if (allInteger) sharedItems.ContainsInteger = true;

                        // Output items only for row / column / filter fields
                        if (pt.RowLabels.Any(p => p.SourceName == xlpf.SourceName)
                            || pt.ColumnLabels.Any(p => p.SourceName == xlpf.SourceName)
                            || pt.ReportFilters.Any(p => p.SourceName == xlpf.SourceName))
                        {
                            foreach (var value in ptfi.DistinctValues)
                                sharedItems.AppendChild(new NumberItem { Val = value.GetNumber() });

                            if (containsBlank) sharedItems.AppendChild(new MissingItem());
                        }

                        sharedItems.MinValue = ptfi.DistinctValues.Min(x => x.GetNumber());
                        sharedItems.MaxValue = ptfi.DistinctValues.Max(x => x.GetNumber());
                    }
                    else if (types.Length == 1 && types.Single() == XLDataType.DateTime)
                    {
                        sharedItems.ContainsSemiMixedTypes = containsBlank;
                        sharedItems.ContainsNonDate = false;
                        sharedItems.ContainsString = false;
                        sharedItems.ContainsDate = true;

                        // Output items only for row / column / filter fields
                        if (pt.RowLabels.Any(p => p.SourceName == xlpf.SourceName)
                            || pt.ColumnLabels.Any(p => p.SourceName == xlpf.SourceName)
                            || pt.ReportFilters.Any(p => p.SourceName == xlpf.SourceName))
                        {
                            foreach (var value in ptfi.DistinctValues)
                                sharedItems.AppendChild(new DateTimeItem { Val = value.GetDateTime() });

                            if (containsBlank) sharedItems.AppendChild(new MissingItem());
                        }

                        sharedItems.MinDate = ptfi.DistinctValues.Min(x => x.GetDateTime());
                        sharedItems.MaxDate = ptfi.DistinctValues.Max(x => x.GetDateTime());
                    }
                    else
                    {
                        if (ptfi.DistinctValues.Any(v => v.GetText().Length > 255))
                            sharedItems.LongText = true;

                        foreach (var value in ptfi.DistinctValues)
                            sharedItems.AppendChild(new StringItem { Val = value.GetText() });

                        if (containsBlank) sharedItems.AppendChild(new MissingItem());
                    }

                    sharedItems.Count = Convert.ToUInt32(sharedItems.Elements().Count());
                }

                if (containsBlank) sharedItems.ContainsBlank = true;

                if (ptfi.IsTotallyBlankField)
                    pti.Fields.Add(xlpf.SourceName, ptfi);
                else if (ptfi.DistinctValues?.Any() ?? false)
                    sharedItems.Count = Convert.ToUInt32(ptfi.DistinctValues.Count());
            }

            // End CacheFields

            var pivotTableCacheRecordsPart = pivotTableCacheDefinitionPart.GetPartsOfType<PivotTableCacheRecordsPart>().Any() ?
                pivotTableCacheDefinitionPart.GetPartsOfType<PivotTableCacheRecordsPart>().First() :
                pivotTableCacheDefinitionPart.AddNewPart<PivotTableCacheRecordsPart>("rId1");

            var pivotCacheRecords = new PivotCacheRecords();
            pivotCacheRecords.AddNamespaceDeclaration("r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            pivotTableCacheRecordsPart.PivotCacheRecords = pivotCacheRecords;

            context.PivotTables.Add(pti.Guid, pti);
        }

        // Generates content of pivotTablePart
        private static void GeneratePivotTablePartContent(
            PivotTablePart pivotTablePart, XLPivotTable pt,
            uint cacheId, SaveContext context)
        {
            var pti = context.PivotTables[(pt as XLPivotTable).Guid];

            var pivotTableDefinition = new PivotTableDefinition
            {
                Name = pt.Name,
                CacheId = cacheId,
                DataCaption = "Values",
                MergeItem = OpenXmlHelper.GetBooleanValue(pt.MergeAndCenterWithLabels, false),
                Indent = Convert.ToUInt32(pt.RowLabelIndent),
                PageOverThenDown = (pt.FilterAreaOrder == XLFilterAreaOrder.OverThenDown),
                PageWrap = Convert.ToUInt32(pt.FilterFieldsPageWrap),
                ShowError = String.IsNullOrEmpty(pt.ErrorValueReplacement),
                UseAutoFormatting = OpenXmlHelper.GetBooleanValue(pt.AutofitColumns, false),
                PreserveFormatting = OpenXmlHelper.GetBooleanValue(pt.PreserveCellFormatting, true),
                RowGrandTotals = OpenXmlHelper.GetBooleanValue(pt.ShowGrandTotalsRows, true),
                ColumnGrandTotals = OpenXmlHelper.GetBooleanValue(pt.ShowGrandTotalsColumns, true),
                SubtotalHiddenItems = OpenXmlHelper.GetBooleanValue(pt.FilteredItemsInSubtotals, false),
                MultipleFieldFilters = OpenXmlHelper.GetBooleanValue(pt.AllowMultipleFilters, true),
                CustomListSort = OpenXmlHelper.GetBooleanValue(pt.UseCustomListsForSorting, true),
                ShowDrill = OpenXmlHelper.GetBooleanValue(pt.ShowExpandCollapseButtons, true),
                ShowDataTips = OpenXmlHelper.GetBooleanValue(pt.ShowContextualTooltips, true),
                ShowMemberPropertyTips = OpenXmlHelper.GetBooleanValue(pt.ShowPropertiesInTooltips, true),
                ShowHeaders = OpenXmlHelper.GetBooleanValue(pt.DisplayCaptionsAndDropdowns, true),
                GridDropZones = OpenXmlHelper.GetBooleanValue(pt.ClassicPivotTableLayout, false),
                ShowEmptyRow = OpenXmlHelper.GetBooleanValue(pt.ShowEmptyItemsOnRows, false),
                ShowEmptyColumn = OpenXmlHelper.GetBooleanValue(pt.ShowEmptyItemsOnColumns, false),
                ShowItems = OpenXmlHelper.GetBooleanValue(pt.DisplayItemLabels, true),
                FieldListSortAscending = OpenXmlHelper.GetBooleanValue(pt.SortFieldsAtoZ, false),
                PrintDrill = OpenXmlHelper.GetBooleanValue(pt.PrintExpandCollapsedButtons, false),
                ItemPrintTitles = OpenXmlHelper.GetBooleanValue(pt.RepeatRowLabels, false),
                FieldPrintTitles = OpenXmlHelper.GetBooleanValue(pt.PrintTitles, false),
                EnableDrill = OpenXmlHelper.GetBooleanValue(pt.EnableShowDetails, true)
            };

            if (!String.IsNullOrEmpty(pt.ColumnHeaderCaption))
                pivotTableDefinition.ColumnHeaderCaption = StringValue.FromString(pt.ColumnHeaderCaption);

            if (!String.IsNullOrEmpty(pt.RowHeaderCaption))
                pivotTableDefinition.RowHeaderCaption = StringValue.FromString(pt.RowHeaderCaption);

            if (pt.ClassicPivotTableLayout)
            {
                pivotTableDefinition.Compact = false;
                pivotTableDefinition.CompactData = false;
            }

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
                FirstHeaderRow = 1U,
                FirstDataRow = 1U,
                FirstDataColumn = 1U
            };

            if (pt.ReportFilters.Any())
            {
                // Reference cell is the part BELOW the report filters
                location.Reference = pt.TargetCell.CellBelow(pt.ReportFilters.Count() + 1).Address.ToString();
            }
            else
                location.Reference = pt.TargetCell.Address.ToString();

            var rowFields = new RowFields();
            var columnFields = new ColumnFields();
            var rowItems = new RowItems();
            var columnItems = new ColumnItems();
            var pageFields = new PageFields { Count = (uint)pt.ReportFilters.Count() };
            var pivotFields = new PivotFields { Count = Convert.ToUInt32(pt.SourceRange.ColumnCount()) };

            var orderedPageFields = new SortedDictionary<int, PageField>();
            var orderedColumnLabels = new SortedDictionary<int, Field>();
            var orderedRowLabels = new SortedDictionary<int, Field>();

            // Add value fields first
            if (pt.Values.Any())
            {
                if (pt.RowLabels.Contains(XLConstants.PivotTable.ValuesSentinalLabel))
                {
                    var f = pt.RowLabels.First(f1 => f1.SourceName == XLConstants.PivotTable.ValuesSentinalLabel);
                    orderedRowLabels.Add(pt.RowLabels.IndexOf(f), new Field { Index = -2 });
                    pivotTableDefinition.DataOnRows = true;
                }
                else if (pt.ColumnLabels.Contains(XLConstants.PivotTable.ValuesSentinalLabel))
                {
                    var f = pt.ColumnLabels.First(f1 => f1.SourceName == XLConstants.PivotTable.ValuesSentinalLabel);
                    orderedColumnLabels.Add(pt.ColumnLabels.IndexOf(f), new Field { Index = -2 });
                }
            }

            // TODO: improve performance as per https://github.com/ClosedXML/ClosedXML/pull/984#discussion_r217266491
            foreach (var xlpf in pt.Fields)
            {
                var ptfi = pti.Fields[xlpf.SourceName];

                if (pt.RowLabels.Contains(xlpf.SourceName))
                {
                    var rowLabelIndex = pt.RowLabels.IndexOf(xlpf);
                    var f = new Field { Index = pt.Fields.IndexOf(xlpf) };
                    orderedRowLabels.Add(rowLabelIndex, f);

                    if (ptfi.IsTotallyBlankField)
                        rowItems.AppendChild(new RowItem());
                    else
                    {
                        for (var i = 0; i < ptfi.DistinctValues.Count(); i++)
                        {
                            var rowItem = new RowItem();
                            rowItem.AppendChild(new MemberPropertyIndex { Val = i });
                            rowItems.AppendChild(rowItem);
                        }
                    }

                    var rowItemTotal = new RowItem { ItemType = ItemValues.Grand };
                    rowItemTotal.AppendChild(new MemberPropertyIndex());
                    rowItems.AppendChild(rowItemTotal);
                }
                else if (pt.ColumnLabels.Contains(xlpf.SourceName))
                {
                    var columnlabelIndex = pt.ColumnLabels.IndexOf(xlpf);
                    var f = new Field { Index = pt.Fields.IndexOf(xlpf) };
                    orderedColumnLabels.Add(columnlabelIndex, f);

                    if (ptfi.IsTotallyBlankField)
                        columnItems.AppendChild(new RowItem());
                    else
                    {
                        for (var i = 0; i < ptfi.DistinctValues.Count(); i++)
                        {
                            var rowItem = new RowItem();
                            rowItem.AppendChild(new MemberPropertyIndex { Val = i });
                            columnItems.AppendChild(rowItem);
                        }
                    }

                    var rowItemTotal = new RowItem { ItemType = ItemValues.Grand };
                    rowItemTotal.AppendChild(new MemberPropertyIndex());
                    columnItems.AppendChild(rowItemTotal);
                }
            }

            foreach (var xlpf in pt.Fields)
            {
                var ptfi = pti.Fields[xlpf.SourceName];
                IXLPivotField labelOrFilterField = null;
                var pf = new PivotField
                {
                    Name = xlpf.CustomName,
                    IncludeNewItemsInFilter = OpenXmlHelper.GetBooleanValue(xlpf.IncludeNewItemsInFilter, false),
                    InsertBlankRow = OpenXmlHelper.GetBooleanValue(xlpf.InsertBlankLines, false),
                    ShowAll = OpenXmlHelper.GetBooleanValue(xlpf.ShowBlankItems, true),
                    InsertPageBreak = OpenXmlHelper.GetBooleanValue(xlpf.InsertPageBreaks, false),
                    AllDrilled = OpenXmlHelper.GetBooleanValue(xlpf.Collapsed, false),
                };
                if (!string.IsNullOrWhiteSpace(xlpf.SubtotalCaption))
                {
                    pf.SubtotalCaption = xlpf.SubtotalCaption;
                }

                if (pt.ClassicPivotTableLayout)
                {
                    pf.Outline = false;
                    pf.Compact = false;
                }
                else
                {
                    pf.Outline = OpenXmlHelper.GetBooleanValue(xlpf.Outline, true);
                    pf.Compact = OpenXmlHelper.GetBooleanValue(xlpf.Compact, true);
                }

                if (xlpf.SortType != XLPivotSortType.Default)
                {
                    pf.SortType = new EnumValue<FieldSortValues>((FieldSortValues)xlpf.SortType);
                }

                switch (pt.Subtotals)
                {
                    case XLPivotSubtotals.DoNotShow:
                        pf.DefaultSubtotal = false;
                        break;

                    case XLPivotSubtotals.AtBottom:
                        pf.SubtotalTop = false;
                        break;

                    case XLPivotSubtotals.AtTop:
                        // at top is by default
                        break;
                }

                if (xlpf.SubtotalsAtTop.HasValue)
                {
                    pf.SubtotalTop = OpenXmlHelper.GetBooleanValue(xlpf.SubtotalsAtTop.Value, true);
                }

                if (pt.RowLabels.Contains(xlpf.SourceName))
                {
                    labelOrFilterField = pt.RowLabels.Get(xlpf.SourceName);
                    pf.Axis = PivotTableAxisValues.AxisRow;
                }
                else if (pt.ColumnLabels.Contains(xlpf.SourceName))
                {
                    labelOrFilterField = pt.ColumnLabels.Get(xlpf.SourceName);
                    pf.Axis = PivotTableAxisValues.AxisColumn;
                }
                else if (pt.ReportFilters.Contains(xlpf.SourceName))
                {
                    labelOrFilterField = pt.ReportFilters.Get(xlpf.SourceName);
                    var sortOrderIndex = pt.ReportFilters.IndexOf(labelOrFilterField);

                    location.ColumnsPerPage = 1;
                    location.RowPageCount = 1;
                    pf.Axis = PivotTableAxisValues.AxisPage;

                    var pageField = new PageField
                    {
                        Hierarchy = -1,
                        Field = pt.Fields.IndexOf(xlpf)
                    };

                    if (labelOrFilterField.SelectedValues.Count == 1)
                    {
                        var selectedValue = labelOrFilterField.SelectedValues.Single();
                        var index = ptfi.DistinctValues.IndexOf(selectedValue, XLCellValueComparer.OrdinalIgnoreCase);
                        if (index >= 0)
                            pageField.Item = (UInt32)index;
                    }

                    orderedPageFields.Add(sortOrderIndex, pageField);
                }

                if ((labelOrFilterField?.SelectedValues?.Count ?? 0) > 1)
                    pf.MultipleItemSelectionAllowed = true;

                if (pt.Values.Any(p => p.SourceName == xlpf.SourceName))
                    pf.DataField = true;

                var fieldItems = new Items();

                // Output items only for row / column / filter fields
                if (!ptfi.IsTotallyBlankField &&
                    ptfi.DistinctValues.Any()
                    && (pt.RowLabels.Contains(xlpf.SourceName)
                        || pt.ColumnLabels.Contains(xlpf.SourceName)
                        || pt.ReportFilters.Contains(xlpf.SourceName)))
                {
                    uint i = 0;
                    foreach (var value in ptfi.DistinctValues)
                    {
                        var item = new Item { Index = i };

                        if (labelOrFilterField != null && labelOrFilterField.Collapsed)
                            item.HideDetails = BooleanValue.FromBoolean(false);

                        if (labelOrFilterField != null &&
                            labelOrFilterField.SelectedValues.Count > 1 &&
                            !labelOrFilterField.SelectedValues.Contains(value, XLCellValueComparer.OrdinalIgnoreCase))
                            item.Hidden = BooleanValue.FromBoolean(true);

                        fieldItems.AppendChild(item);

                        i++;
                    }
                }

                if (xlpf.Subtotals.Any())
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
                // If the field itself doesn't have subtotals, but the pivot table is set to show pivot tables, add the default item
                else if (pt.Subtotals != XLPivotSubtotals.DoNotShow)
                {
                    fieldItems.AppendChild(new Item { ItemType = ItemValues.Default });
                }

                if (fieldItems.Any())
                {
                    fieldItems.Count = Convert.ToUInt32(fieldItems.Count());
                    pf.AppendChild(fieldItems);
                }

                #region Excel 2010 Features

                if (xlpf.RepeatItemLabels)
                {
                    var pivotFieldExtensionList = new PivotFieldExtensionList();
                    pivotFieldExtensionList.RemoveNamespaceDeclaration("x");
                    var pivotFieldExtension = new PivotFieldExtension { Uri = "{2946ED86-A175-432a-8AC1-64E0C546D7DE}" };
                    pivotFieldExtension.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");

                    var pivotField2 = new DocumentFormat.OpenXml.Office2010.Excel.PivotField { FillDownLabels = true };

                    pivotFieldExtension.AppendChild(pivotField2);

                    pivotFieldExtensionList.AppendChild(pivotFieldExtension);
                    pf.AppendChild(pivotFieldExtensionList);
                }

                #endregion Excel 2010 Features

                pivotFields.AppendChild(pf);
            }

            pivotTableDefinition.AppendChild(location);
            pivotTableDefinition.AppendChild(pivotFields);

            if (pt.RowLabels.Any())
            {
                rowFields.Append(orderedRowLabels.Values);
                rowFields.Count = Convert.ToUInt32(rowFields.Count());
                pivotTableDefinition.AppendChild(rowFields);
            }
            else
            {
                rowItems.AppendChild(new RowItem());
            }

            if (rowItems.Any())
            {
                rowItems.Count = Convert.ToUInt32(rowItems.Count());
                pivotTableDefinition.AppendChild(rowItems);
            }

            if (pt.ColumnLabels.All(cl => cl.CustomName == XLConstants.PivotTable.ValuesSentinalLabel))
            {
                for (int i = 0; i < pt.Values.Count(); i++)
                {
                    var rowItem = new RowItem();
                    rowItem.Index = Convert.ToUInt32(i);
                    rowItem.AppendChild(new MemberPropertyIndex() { Val = i });
                    columnItems.AppendChild(rowItem);
                }
            }

            if (pt.ColumnLabels.Any())
            {
                columnFields.Append(orderedColumnLabels.Values);
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
                pageFields.Append(orderedPageFields.Values);
                pageFields.Count = Convert.ToUInt32(pageFields.Count());
                pivotTableDefinition.AppendChild(pageFields);
            }

            var dataFields = new DataFields();
            foreach (var value in pt.Values)
            {
                var sourceColumn =
                    pt.SourceRange.Columns().FirstOrDefault(c => c.Cell(1).Value.ToString(CultureInfo.CurrentCulture) == value.SourceName);
                if (sourceColumn == null) continue;

                UInt32 numberFormatId = 0;
                if (value.NumberFormat.NumberFormatId != -1 || context.SharedNumberFormats.ContainsKey(value.NumberFormat.NumberFormatId))
                    numberFormatId = (UInt32)value.NumberFormat.NumberFormatId;
                else if (context.SharedNumberFormats.Any(snf => snf.Value.NumberFormat.Format == value.NumberFormat.Format))
                    numberFormatId = (UInt32)context.SharedNumberFormats.First(snf => snf.Value.NumberFormat.Format == value.NumberFormat.Format).Key;

                var df = new DataField
                {
                    Name = value.CustomName,
                    Field = (UInt32)(sourceColumn.ColumnNumber() - pt.SourceRange.RangeAddress.FirstAddress.ColumnNumber),
                    Subtotal = value.SummaryFormula.ToOpenXml(),
                    ShowDataAs = value.Calculation.ToOpenXml(),
                    NumberFormatId = numberFormatId
                };

                if (!String.IsNullOrEmpty(value.BaseField))
                {
                    var baseField = pt.SourceRange.Columns().FirstOrDefault(c => c.Cell(1).Value.TryGetText(out var headerText) && headerText == value.BaseField);
                    if (baseField != null)
                    {
                        df.BaseField = baseField.ColumnNumber() - pt.SourceRange.RangeAddress.FirstAddress.ColumnNumber;

                        var items = baseField.CellsUsed()
                            .Select(c => c.Value)
                            .Skip(1) // Skip header column
                            .Distinct().ToList();

                        var indexOfItem = items.IndexOf(value.BaseItem);
                        if (indexOfItem >= 0)
                            df.BaseItem = Convert.ToUInt32(indexOfItem);
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

            if (dataFields.Any())
            {
                dataFields.Count = Convert.ToUInt32(dataFields.Count());
                pivotTableDefinition.AppendChild(dataFields);
            }

            var pts = new PivotTableStyle
            {
                ShowRowHeaders = pt.ShowRowHeaders,
                ShowColumnHeaders = pt.ShowColumnHeaders,
                ShowRowStripes = pt.ShowRowStripes,
                ShowColumnStripes = pt.ShowColumnStripes
            };

            if (pt.Theme != XLPivotTableTheme.None)
                pts.Name = Enum.GetName(typeof(XLPivotTableTheme), pt.Theme);

            pivotTableDefinition.AppendChild(pts);

            // Pivot formats
            if (pivotTableDefinition.Formats == null)
                pivotTableDefinition.Formats = new Formats();
            else
                pivotTableDefinition.Formats.RemoveAllChildren();

            foreach (var styleFormat in pt.StyleFormats.RowGrandTotalFormats)
                GeneratePivotTableFormat(isRow: true, (XLPivotStyleFormat)styleFormat, pivotTableDefinition, context);

            foreach (var styleFormat in pt.StyleFormats.ColumnGrandTotalFormats)
                GeneratePivotTableFormat(isRow: false, (XLPivotStyleFormat)styleFormat, pivotTableDefinition, context);

            foreach (var pivotField in pt.ImplementedFields)
            {
                GeneratePivotFieldFormat(XLPivotStyleFormatTarget.Header, pt, (XLPivotField)pivotField, (XLPivotStyleFormat)pivotField.StyleFormats.Header, pivotTableDefinition, context);
                GeneratePivotFieldFormat(XLPivotStyleFormatTarget.Subtotal, pt, (XLPivotField)pivotField, (XLPivotStyleFormat)pivotField.StyleFormats.Subtotal, pivotTableDefinition, context);
                GeneratePivotFieldFormat(XLPivotStyleFormatTarget.Label, pt, (XLPivotField)pivotField, (XLPivotStyleFormat)pivotField.StyleFormats.Label, pivotTableDefinition, context);
                GeneratePivotFieldFormat(XLPivotStyleFormatTarget.Data, pt, (XLPivotField)pivotField, (XLPivotStyleFormat)pivotField.StyleFormats.DataValuesFormat, pivotTableDefinition, context);
            }

            if (pivotTableDefinition.Formats.Any())
            {
                pivotTableDefinition.Formats.Count = new UInt32Value((uint)pivotTableDefinition.Formats.Count());
            }
            else
                pivotTableDefinition.Formats = null;

            #region Excel 2010 Features

            var pivotTableDefinitionExtensionList = new PivotTableDefinitionExtensionList();

            var pivotTableDefinitionExtension = new PivotTableDefinitionExtension { Uri = "{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}" };
            pivotTableDefinitionExtension.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");

            var pivotTableDefinition2 = new DocumentFormat.OpenXml.Office2010.Excel.PivotTableDefinition
            {
                EnableEdit = pt.EnableCellEditing,
                HideValuesRow = !pt.ShowValuesRow
            };
            pivotTableDefinition2.AddNamespaceDeclaration("xm", "http://schemas.microsoft.com/office/excel/2006/main");

            pivotTableDefinitionExtension.AppendChild(pivotTableDefinition2);

            pivotTableDefinitionExtensionList.AppendChild(pivotTableDefinitionExtension);
            pivotTableDefinition.AppendChild(pivotTableDefinitionExtensionList);

            #endregion Excel 2010 Features

            pivotTablePart.PivotTableDefinition = pivotTableDefinition;
        }

        private static void GeneratePivotTableFormat(Boolean isRow, XLPivotStyleFormat styleFormat, PivotTableDefinition pivotTableDefinition, SaveContext context)
        {
            if (DefaultStyle.Equals(styleFormat.Style) || !context.DifferentialFormats.ContainsKey(((XLStyle)styleFormat.Style).Value))
                return;

            var format = new Format();

            format.FormatId = UInt32Value.FromUInt32(Convert.ToUInt32(context.DifferentialFormats[((XLStyle)styleFormat.Style).Value]));

            var pivotArea = GenerateDefaultPivotArea(XLPivotStyleFormatTarget.GrandTotal);

            pivotArea.LabelOnly = OpenXmlHelper.GetBooleanValue(styleFormat.AppliesTo == XLPivotStyleFormatElement.Label, false);
            pivotArea.DataOnly = OpenXmlHelper.GetBooleanValue(styleFormat.AppliesTo == XLPivotStyleFormatElement.Data, true);

            pivotArea.GrandColumn = OpenXmlHelper.GetBooleanValue(!isRow, false);
            pivotArea.GrandRow = OpenXmlHelper.GetBooleanValue(isRow, false);
            pivotArea.Axis = isRow ? PivotTableAxisValues.AxisRow : PivotTableAxisValues.AxisColumn;

            format.PivotArea = pivotArea;

            pivotTableDefinition.Formats.AppendChild(format);
        }

        private static void GeneratePivotFieldFormat(XLPivotStyleFormatTarget target, XLPivotTable pt, XLPivotField pivotField, XLPivotStyleFormat styleFormat, PivotTableDefinition pivotTableDefinition, SaveContext context)
        {
            if (target == XLPivotStyleFormatTarget.GrandTotal)
                throw new ArgumentException($"Use {nameof(GeneratePivotTableFormat)} to populate grand total formats.");

            if (DefaultStyle.Equals(styleFormat.Style) || !context.DifferentialFormats.ContainsKey(((XLStyle)styleFormat.Style).Value))
                return;

            var format = new Format();

            format.FormatId = UInt32Value.FromUInt32(Convert.ToUInt32(context.DifferentialFormats[((XLStyle)styleFormat.Style).Value]));

            var pivotArea = GenerateDefaultPivotArea(target);

            pivotArea.LabelOnly = OpenXmlHelper.GetBooleanValue(styleFormat.AppliesTo == XLPivotStyleFormatElement.Label, false);
            pivotArea.DataOnly = OpenXmlHelper.GetBooleanValue(styleFormat.AppliesTo == XLPivotStyleFormatElement.Data, true);

            pivotArea.CollapsedLevelsAreSubtotals = OpenXmlHelper.GetBooleanValue(styleFormat.CollapsedLevelsAreSubtotals, false);

            if (target == XLPivotStyleFormatTarget.Header)
            {
                pivotArea.Field = pivotField.Offset;

                if (pivotField.IsOnRowAxis)
                    pivotArea.Axis = PivotTableAxisValues.AxisRow;
                else if (pivotField.IsOnColumnAxis)
                    pivotArea.Axis = PivotTableAxisValues.AxisColumn;
                else if (pivotField.IsInFilterList)
                    pivotArea.Axis = PivotTableAxisValues.AxisPage;
                else
                    throw new NotImplementedException();
            }

            //Ensure referenced pivot field is added to field references
            if (new[]
                {
                    XLPivotStyleFormatTarget.Data, XLPivotStyleFormatTarget.Label, XLPivotStyleFormatTarget.Subtotal
                }.Contains(target)
                && !styleFormat.FieldReferences.OfType<PivotLabelFieldReference>().Select(fr => fr.PivotField).Contains(pivotField))
            {
                var fr = new PivotLabelFieldReference(pivotField);
                fr.DefaultSubtotal = target == XLPivotStyleFormatTarget.Subtotal;
                styleFormat.FieldReferences.Insert(0, fr);
            }

            if (pivotArea.PivotAreaReferences == null)
                pivotArea.PivotAreaReferences = new PivotAreaReferences();
            else
                pivotArea.PivotAreaReferences.RemoveAllChildren();

            foreach (var fr in styleFormat.FieldReferences)
            {
                GeneratePivotAreaReference(pt, pivotArea.PivotAreaReferences, fr, context);
            }

            if (pivotArea.PivotAreaReferences.Any())
            {
                pivotArea.PivotAreaReferences.Count = new UInt32Value((uint)pivotArea.PivotAreaReferences.Count());
            }
            else
                pivotArea.PivotAreaReferences = null;

            format.PivotArea = pivotArea;
            pivotTableDefinition.Formats.AppendChild(format);
        }

        private static PivotArea GenerateDefaultPivotArea(XLPivotStyleFormatTarget target)
        {
            switch (target)
            {
                case XLPivotStyleFormatTarget.Header:
                    return new PivotArea
                    {
                        Type = PivotAreaValues.Button,
                        FieldPosition = 0,
                        DataOnly = OpenXmlHelper.GetBooleanValue(false, true),
                        LabelOnly = OpenXmlHelper.GetBooleanValue(true, false),
                        Outline = OpenXmlHelper.GetBooleanValue(false, true),
                    };

                case XLPivotStyleFormatTarget.Subtotal:
                    return new PivotArea
                    {
                        Type = PivotAreaValues.Normal,
                        FieldPosition = 0,
                    };

                case XLPivotStyleFormatTarget.GrandTotal:
                    return new PivotArea
                    {
                        Type = PivotAreaValues.Normal,
                        FieldPosition = 0,
                        DataOnly = OpenXmlHelper.GetBooleanValue(false, true),
                        LabelOnly = OpenXmlHelper.GetBooleanValue(false, false),
                    };

                case XLPivotStyleFormatTarget.Label:
                    return new PivotArea
                    {
                        Type = PivotAreaValues.Normal,
                        FieldPosition = 0,
                        DataOnly = OpenXmlHelper.GetBooleanValue(false, true),
                        LabelOnly = OpenXmlHelper.GetBooleanValue(true, false),
                    };

                case XLPivotStyleFormatTarget.Data:
                    return new PivotArea
                    {
                        Type = PivotAreaValues.Normal,
                        FieldPosition = 0,
                    };

                default:
                    throw new NotImplementedException();
            }
        }

        private static void GeneratePivotAreaReference(XLPivotTable pt, PivotAreaReferences pivotAreaReferences, AbstractPivotFieldReference fieldReference, SaveContext context)
        {
            var pivotAreaReference = new PivotAreaReference();

            pivotAreaReference.DefaultSubtotal = OpenXmlHelper.GetBooleanValue(fieldReference.DefaultSubtotal, false);
            pivotAreaReference.Field = fieldReference.GetFieldOffset();

            var matchedOffsets = fieldReference.Match(context.PivotTables[pt.Guid], pt);
            foreach (var o in matchedOffsets)
            {
                pivotAreaReference.AppendChild(new FieldItem { Val = UInt32Value.FromUInt32((uint)o) });
            }

            pivotAreaReferences.AppendChild(pivotAreaReference);
        }

        private static void GenerateWorksheetCommentsPartContent(WorksheetCommentsPart worksheetCommentsPart,
            XLWorksheet xlWorksheet)
        {
            var commentList = new CommentList();
            var authorsDict = new Dictionary<String, Int32>();
            foreach (var c in xlWorksheet.Internals.CellsCollection.GetCells(c => c.HasComment))
            {
                var comment = new Comment { Reference = c.Address.ToStringRelative() };
                var authorName = c.GetComment().Author;

                if (!authorsDict.TryGetValue(authorName, out int authorId))
                {
                    authorId = authorsDict.Count;
                    authorsDict.Add(authorName, authorId);
                }
                comment.AuthorId = (UInt32)authorId;

                var commentText = new CommentText();
                foreach (var rt in c.GetComment())
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

            worksheetCommentsPart.Comments.Append(authors);
            worksheetCommentsPart.Comments.Append(commentList);
        }

        // Generates content of vmlDrawingPart1.
        private static bool GenerateVmlDrawingPartContent(VmlDrawingPart vmlDrawingPart, XLWorksheet xlWorksheet)
        {
            using (var ms = new MemoryStream())
            using (var stream = vmlDrawingPart.GetStream(FileMode.OpenOrCreate))
            {
                CopyStream(stream, ms);
                stream.Position = 0;
                var writer = new XmlTextWriter(stream, Encoding.UTF8);

                writer.WriteStartElement("xml");

                // https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.vml.shapetype?view=openxml-2.8.1#remarks
                // This element defines a shape template that can be used to create other shapes.
                // Shapetype is identical to the shape element(§14.1.2.19) except it cannot reference another shapetype element.
                // The type attribute shall not be used with shapetype.
                // Attributes defined in the shape override any that appear in the shapetype positioning attributes
                // (such as top, width, z-index, rotation, flip) are not passed to a shape from a shapetype.
                // To use this element, create a shapetype with a specific id attribute.
                // Then create a shape and reference the shapetype's id using the type attribute.
                new Vml.Shapetype(
                    new Vml.Stroke { JoinStyle = Vml.StrokeJoinStyleValues.Miter },
                    new Vml.Path { AllowGradientShape = true, ConnectionPointType = ConnectValues.Rectangle }
                    )
                {
                    Id = XLConstants.Comment.ShapeTypeId,
                    CoordinateSize = "21600,21600",
                    OptionalNumber = 202,
                    EdgePath = "m,l,21600r21600,l21600,xe",
                }
                    .WriteTo(writer);

                var cellWithComments = xlWorksheet.Internals.CellsCollection.GetCells(c => c.HasComment);

                var hasAnyVmlElements = false;

                foreach (var c in cellWithComments)
                {
                    GenerateCommentShape(c).WriteTo(writer);
                    hasAnyVmlElements |= true;
                }

                if (ms.Length > 0)
                {
                    ms.Position = 0;
                    var xdoc = XDocumentExtensions.Load(ms);
                    xdoc.Root.Elements().ForEach(e => writer.WriteRaw(e.ToString()));
                    hasAnyVmlElements |= xdoc.Root.HasElements;
                }

                writer.WriteEndElement();
                writer.Flush();
                writer.Close();

                return hasAnyVmlElements;
            }
        }

        // VML Shape for Comment
        private static Vml.Shape GenerateCommentShape(XLCell c)
        {
            var rowNumber = c.Address.RowNumber;
            var columnNumber = c.Address.ColumnNumber;

            var comment = c.GetComment();
            var shapeId = String.Concat("_x0000_s", comment.ShapeId);
            // Unique per cell (workbook?), e.g.: "_x0000_s1026"
            var anchor = GetAnchor(c);
            var textBox = GetTextBox(comment.Style);
            var fill = new Vml.Fill { Color2 = "#" + comment.Style.ColorsAndLines.FillColor.Color.ToHex().Substring(2) };
            if (comment.Style.ColorsAndLines.FillTransparency < 1)
                fill.Opacity =
                    Math.Round(Convert.ToDouble(comment.Style.ColorsAndLines.FillTransparency), 2).ToInvariantString();
            var stroke = GetStroke(c);
            var shape = new Vml.Shape(
                fill,
                stroke,
                new Vml.Shadow { Color = "black", Obscured = true },
                new Vml.Path { ConnectionPointType = ConnectValues.None },
                textBox,
                new ClientData(
                    new MoveWithCells(comment.Style.Properties.Positioning == XLDrawingAnchor.Absolute
                        ? "True"
                        : "False"), // Counterintuitive
                    new ResizeWithCells(comment.Style.Properties.Positioning == XLDrawingAnchor.MoveAndSizeWithCells
                        ? "False"
                        : "True"), // Counterintuitive
                    anchor,
                    new HorizontalTextAlignment(comment.Style.Alignment.Horizontal.ToString().ToCamel()),
                    new Vml.Spreadsheet.VerticalTextAlignment(comment.Style.Alignment.Vertical.ToString().ToCamel()),
                    new AutoFill("False"),
                    new CommentRowTarget { Text = (rowNumber - 1).ToInvariantString() },
                    new CommentColumnTarget { Text = (columnNumber - 1).ToInvariantString() },
                    new Locked(comment.Style.Protection.Locked ? "True" : "False"),
                    new LockText(comment.Style.Protection.LockText ? "True" : "False"),
                    new Visible(comment.Visible ? "True" : "False")
                    )
                { ObjectType = ObjectValues.Note }
                )
            {
                Id = shapeId,
                Type = "#" + XLConstants.Comment.ShapeTypeId,
                Style = GetCommentStyle(c),
                FillColor = "#" + comment.Style.ColorsAndLines.FillColor.Color.ToHex().Substring(2),
                StrokeColor = "#" + comment.Style.ColorsAndLines.LineColor.Color.ToHex().Substring(2),
                StrokeWeight = String.Concat(comment.Style.ColorsAndLines.LineWeight.ToInvariantString(), "pt"),
                InsetMode = comment.Style.Margins.Automatic ? InsetMarginValues.Auto : InsetMarginValues.Custom
            };
            if (!String.IsNullOrWhiteSpace(comment.Style.Web.AlternateText))
                shape.Alternate = comment.Style.Web.AlternateText;

            return shape;
        }

        private static Vml.Stroke GetStroke(XLCell c)
        {
            var lineDash = c.GetComment().Style.ColorsAndLines.LineDash;
            var stroke = new Vml.Stroke
            {
                LineStyle = c.GetComment().Style.ColorsAndLines.LineStyle.ToOpenXml(),
                DashStyle =
                    lineDash == XLDashStyle.RoundDot || lineDash == XLDashStyle.SquareDot
                        ? "shortDot"
                        : lineDash.ToString().ToCamel()
            };
            if (lineDash == XLDashStyle.RoundDot)
                stroke.EndCap = Vml.StrokeEndCapValues.Round;
            if (c.GetComment().Style.ColorsAndLines.LineTransparency < 1)
                stroke.Opacity =
                    Math.Round(Convert.ToDouble(c.GetComment().Style.ColorsAndLines.LineTransparency), 2).ToInvariantString();
            return stroke;
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

            var tb = new Vml.TextBox();

            if (sb.Length > 0)
                tb.Style = sb.ToString();

            var dm = ds.Margins;
            if (!dm.Automatic)
                tb.Inset = String.Concat(
                    dm.Left.ToInvariantString(), "in,",
                    dm.Top.ToInvariantString(), "in,",
                    dm.Right.ToInvariantString(), "in,",
                    dm.Bottom.ToInvariantString(), "in");

            return tb;
        }

        private static Anchor GetAnchor(XLCell cell)
        {
            var c = cell.GetComment();
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
                Text = string.Concat(
                    fcNumber, ", ", fcOffset, ", ",
                    frNumber, ", ", frOffset, ", ",
                    lcNumber, ", ", lcOffset, ", ",
                    lrNumber, ", ", lrOffset
                    )
            };
        }

        private static StringValue GetCommentStyle(XLCell cell)
        {
            var c = cell.GetComment();
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
            sb.Append(c.ZOrder.ToInvariantString());

            return sb.ToString();
        }

        #region GenerateWorkbookStylesPartContent

        private void GenerateWorkbookStylesPartContent(WorkbookStylesPart workbookStylesPart, SaveContext context)
        {
            var defaultStyle = DefaultStyleValue;

            if (!context.SharedFonts.ContainsKey(defaultStyle.Font))
                context.SharedFonts.Add(defaultStyle.Font, new FontInfo { FontId = 0, Font = defaultStyle.Font });

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

            context.SharedStyles.Add(defaultStyle,
                new StyleInfo
                {
                    StyleId = defaultFormatId,
                    Style = defaultStyle,
                    FontId = 0,
                    FillId = 0,
                    BorderId = 0,
                    IncludeQuotePrefix = false,
                    NumberFormatId = 0
                    //AlignmentId = 0
                });

            UInt32 styleCount = 1;
            UInt32 fontCount = 1;
            UInt32 fillCount = 3;
            UInt32 borderCount = 1;
            var numberFormatCount = 0; // 0-based
            var pivotTableNumberFormats = new HashSet<IXLPivotValueFormat>();
            var xlStyles = new HashSet<XLStyleValue>();

            foreach (var worksheet in WorksheetsInternal)
            {
                xlStyles.Add(worksheet.StyleValue);
                foreach (var s in worksheet.Internals.ColumnsCollection.Select(c => c.Value.StyleValue))
                {
                    xlStyles.Add(s);
                }
                foreach (var s in worksheet.Internals.RowsCollection.Select(r => r.Value.StyleValue))
                {
                    xlStyles.Add(s);
                }

                foreach (var s in worksheet.Internals.CellsCollection.GetCells().Select(c => c.StyleValue))
                {
                    xlStyles.Add(s);
                }

                foreach (var ptnf in worksheet.PivotTables.SelectMany(pt => pt.Values.Select(ptv => ptv.NumberFormat)).Distinct().Where(nf => !pivotTableNumberFormats.Contains(nf)))
                    pivotTableNumberFormats.Add(ptnf);
            }

            var alignments = xlStyles.Select(s => s.Alignment).Distinct().ToList();
            var borders = xlStyles.Select(s => s.Border).Distinct().ToList();
            var fonts = xlStyles.Select(s => s.Font).Distinct().ToList();
            var fills = xlStyles.Select(s => s.Fill).Distinct().ToList();
            var numberFormats = xlStyles.Select(s => s.NumberFormat).Distinct().ToList();
            var protections = xlStyles.Select(s => s.Protection).Distinct().ToList();

            for (int i = 0; i < fonts.Count; i++)
            {
                if (!context.SharedFonts.ContainsKey(fonts[i]))
                {
                    context.SharedFonts.Add(fonts[i], new FontInfo { FontId = (uint)fontCount++, Font = fonts[i] });
                }
            }

            var sharedFills = fills.ToDictionary(
                f => f, f => new FillInfo { FillId = fillCount++, Fill = f });

            var sharedBorders = borders.ToDictionary(
                b => b, b => new BorderInfo { BorderId = borderCount++, Border = b });

            var sharedNumberFormats = numberFormats
                .Where(nf => nf.NumberFormatId == -1)
                .ToDictionary(nf => nf, nf => new NumberFormatInfo
                {
                    NumberFormatId = XLConstants.NumberOfBuiltInStyles + numberFormatCount++,
                    NumberFormat = nf
                });

            foreach (var pivotNumberFormat in pivotTableNumberFormats.Where(nf => nf.NumberFormatId == -1))
            {
                var numberFormatKey = new XLNumberFormatKey
                {
                    NumberFormatId = -1,
                    Format = pivotNumberFormat.Format
                };
                var numberFormat = XLNumberFormatValue.FromKey(ref numberFormatKey);

                if (sharedNumberFormats.ContainsKey(numberFormat))
                    continue;

                sharedNumberFormats.Add(numberFormat,
                    new NumberFormatInfo
                    {
                        NumberFormatId = XLConstants.NumberOfBuiltInStyles + numberFormatCount++,
                        NumberFormat = numberFormat
                    });
            }

            var allSharedNumberFormats = ResolveNumberFormats(workbookStylesPart, sharedNumberFormats, defaultFormatId);
            foreach (var nf in allSharedNumberFormats)
            {
                context.SharedNumberFormats.Add(nf.Value.NumberFormatId, nf.Value);
            }

            ResolveFonts(workbookStylesPart, context);
            var allSharedFills = ResolveFills(workbookStylesPart, sharedFills);
            var allSharedBorders = ResolveBorders(workbookStylesPart, sharedBorders);

            foreach (var xlStyle in xlStyles)
            {
                var numberFormatId = xlStyle.NumberFormat.NumberFormatId >= 0
                    ? xlStyle.NumberFormat.NumberFormatId
                    : allSharedNumberFormats[xlStyle.NumberFormat].NumberFormatId;

                if (!context.SharedStyles.ContainsKey(xlStyle))
                    context.SharedStyles.Add(xlStyle,
                        new StyleInfo
                        {
                            StyleId = styleCount++,
                            Style = xlStyle,
                            FontId = context.SharedFonts[xlStyle.Font].FontId,
                            FillId = allSharedFills[xlStyle.Fill].FillId,
                            BorderId = allSharedBorders[xlStyle.Border].BorderId,
                            NumberFormatId = numberFormatId,
                            IncludeQuotePrefix = xlStyle.IncludeQuotePrefix
                        });
            }

            ResolveCellStyleFormats(workbookStylesPart, context);
            ResolveRest(workbookStylesPart, context);

            if (!workbookStylesPart.Stylesheet.CellStyles.Elements<CellStyle>().Any(c => c.BuiltinId != null && c.BuiltinId.HasValue && c.BuiltinId.Value == 0U))
                workbookStylesPart.Stylesheet.CellStyles.AppendChild(new CellStyle { Name = "Normal", FormatId = defaultFormatId, BuiltinId = 0U });

            workbookStylesPart.Stylesheet.CellStyles.Count = (UInt32)workbookStylesPart.Stylesheet.CellStyles.Count();

            var newSharedStyles = new Dictionary<XLStyleValue, StyleInfo>();
            foreach (var ss in context.SharedStyles)
            {
                var styleId = -1;
                foreach (CellFormat f in workbookStylesPart.Stylesheet.CellFormats)
                {
                    styleId++;
                    if (CellFormatsAreEqual(f, ss.Value, compareAlignment: true))
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

        /// <summary>
        /// Populates the differential formats that are currently in the file to the SaveContext
        /// </summary>
        /// <param name="workbookStylesPart">The workbook styles part.</param>
        /// <param name="context">The context.</param>
        private void AddDifferentialFormats(WorkbookStylesPart workbookStylesPart, SaveContext context)
        {
            if (workbookStylesPart.Stylesheet.DifferentialFormats == null)
                workbookStylesPart.Stylesheet.DifferentialFormats = new DifferentialFormats();

            var differentialFormats = workbookStylesPart.Stylesheet.DifferentialFormats;
            differentialFormats.RemoveAllChildren();
            FillDifferentialFormatsCollection(differentialFormats, context.DifferentialFormats);

            foreach (var ws in Worksheets)
            {
                foreach (var cf in ws.ConditionalFormats)
                {
                    var styleValue = (cf.Style as XLStyle).Value;
                    if (!styleValue.Equals(DefaultStyleValue) && !context.DifferentialFormats.ContainsKey(styleValue))
                        AddConditionalDifferentialFormat(workbookStylesPart.Stylesheet.DifferentialFormats, cf, context);
                }

                foreach (var tf in ws.Tables.SelectMany(t => t.Fields))
                {
                    if (tf.IsConsistentStyle())
                    {
                        var style = (tf.Column.Cells()
                            .Skip(tf.Table.ShowHeaderRow ? 1 : 0)
                            .First()
                            .Style as XLStyle).Value;

                        if (!style.Equals(DefaultStyleValue) && !context.DifferentialFormats.ContainsKey(style))
                            AddStyleAsDifferentialFormat(workbookStylesPart.Stylesheet.DifferentialFormats, style, context);
                    }
                }

                foreach (var pt in ws.PivotTables.Cast<XLPivotTable>())
                {
                    foreach (var styleFormat in pt.AllStyleFormats)
                    {
                        var xlStyle = (XLStyle)styleFormat.Style;
                        if (!xlStyle.Value.Equals(DefaultStyleValue) && !context.DifferentialFormats.ContainsKey(xlStyle.Value))
                            AddStyleAsDifferentialFormat(workbookStylesPart.Stylesheet.DifferentialFormats, xlStyle.Value, context);
                    }
                }
            }

            differentialFormats.Count = (UInt32)differentialFormats.Count();
            if (differentialFormats.Count == 0)
                workbookStylesPart.Stylesheet.DifferentialFormats = null;
        }

        private void FillDifferentialFormatsCollection(DifferentialFormats differentialFormats,
            Dictionary<XLStyleValue, int> dictionary)
        {
            dictionary.Clear();
            var id = 0;

            foreach (var df in differentialFormats.Elements<DifferentialFormat>())
            {
                var emptyContainer = new XLStylizedEmpty(DefaultStyle);

                var style = new XLStyle(emptyContainer, DefaultStyle);
                LoadFont(df.Font, emptyContainer.Style.Font);
                LoadBorder(df.Border, emptyContainer.Style.Border);
                LoadNumberFormat(df.NumberingFormat, emptyContainer.Style.NumberFormat);
                LoadFill(df.Fill, emptyContainer.Style.Fill, differentialFillFormat: true);

                if (!dictionary.ContainsKey(emptyContainer.StyleValue))
                    dictionary.Add(emptyContainer.StyleValue, id++);
            }
        }

        private static void AddConditionalDifferentialFormat(DifferentialFormats differentialFormats, IXLConditionalFormat cf,
            SaveContext context)
        {
            var differentialFormat = new DifferentialFormat();
            var styleValue = (cf.Style as XLStyle).Value;

            var diffFont = GetNewFont(new FontInfo { Font = styleValue.Font }, false);
            if (diffFont?.HasChildren ?? false)
                differentialFormat.Append(diffFont);

            if (!String.IsNullOrWhiteSpace(cf.Style.NumberFormat.Format))
            {
                var numberFormat = new NumberingFormat
                {
                    NumberFormatId = (UInt32)(XLConstants.NumberOfBuiltInStyles + differentialFormats.Count()),
                    FormatCode = cf.Style.NumberFormat.Format
                };
                differentialFormat.Append(numberFormat);
            }

            var diffFill = GetNewFill(new FillInfo { Fill = styleValue.Fill }, differentialFillFormat: true, ignoreMod: false);
            if (diffFill?.HasChildren ?? false)
                differentialFormat.Append(diffFill);

            var diffBorder = GetNewBorder(new BorderInfo { Border = styleValue.Border }, false);
            if (diffBorder?.HasChildren ?? false)
                differentialFormat.Append(diffBorder);

            differentialFormats.Append(differentialFormat);

            context.DifferentialFormats.Add(styleValue, differentialFormats.Count() - 1);
        }

        private static void AddStyleAsDifferentialFormat(DifferentialFormats differentialFormats, XLStyleValue style,
            SaveContext context)
        {
            var differentialFormat = new DifferentialFormat();

            var diffFont = GetNewFont(new FontInfo { Font = style.Font }, false);
            if (diffFont?.HasChildren ?? false)
                differentialFormat.Append(diffFont);

            if (!String.IsNullOrWhiteSpace(style.NumberFormat.Format) || style.NumberFormat.NumberFormatId != 0)
            {
                var numberFormat = new NumberingFormat();

                if (style.NumberFormat.NumberFormatId == -1)
                {
                    numberFormat.FormatCode = style.NumberFormat.Format;
                    numberFormat.NumberFormatId = (UInt32)(XLConstants.NumberOfBuiltInStyles +
                        differentialFormats
                            .Descendants<DifferentialFormat>()
                            .Count(df => df.NumberingFormat != null && df.NumberingFormat.NumberFormatId != null && df.NumberingFormat.NumberFormatId.Value >= XLConstants.NumberOfBuiltInStyles));
                }
                else
                {
                    numberFormat.NumberFormatId = (UInt32)(style.NumberFormat.NumberFormatId);
                    if (!string.IsNullOrEmpty(style.NumberFormat.Format))
                        numberFormat.FormatCode = style.NumberFormat.Format;
                    else if (XLPredefinedFormat.FormatCodes.TryGetValue(style.NumberFormat.NumberFormatId, out string formatCode))
                        numberFormat.FormatCode = formatCode;
                }

                differentialFormat.Append(numberFormat);
            }

            var diffFill = GetNewFill(new FillInfo { Fill = style.Fill }, differentialFillFormat: true, ignoreMod: false);
            if (diffFill?.HasChildren ?? false)
                differentialFormat.Append(diffFill);

            var diffBorder = GetNewBorder(new BorderInfo { Border = style.Border }, false);
            if (diffBorder?.HasChildren ?? false)
                differentialFormat.Append(diffBorder);

            differentialFormats.Append(differentialFormat);

            context.DifferentialFormats.Add(style, differentialFormats.Count() - 1);
        }

        private static void ResolveRest(WorkbookStylesPart workbookStylesPart, SaveContext context)
        {
            if (workbookStylesPart.Stylesheet.CellFormats == null)
                workbookStylesPart.Stylesheet.CellFormats = new CellFormats();

            foreach (var styleInfo in context.SharedStyles.Values)
            {
                var info = styleInfo;
                var foundOne =
                    workbookStylesPart.Stylesheet.CellFormats.Cast<CellFormat>().Any(f => CellFormatsAreEqual(f, info, compareAlignment: true));

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
                        f => CellFormatsAreEqual(f, info, compareAlignment: false));

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
                QuotePrefix = OpenXmlHelper.GetBooleanValue(styleInfo.IncludeQuotePrefix, false),
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

        /// <summary>
        /// Check if two style are equivalent.
        /// </summary>
        /// <param name="f">Style in the OpenXML format.</param>
        /// <param name="styleInfo">Style in the ClosedXML format.</param>
        /// <param name="compareAlignment">Flag specifying whether or not compare the alignments of two styles.
        /// Styles in x:cellStyleXfs section do not include alignment so we don't have to compare it in this case.
        /// Styles in x:cellXfs section, on the opposite, do include alignments, and we must compare them.
        /// </param>
        /// <returns>True if two formats are equivalent, false otherwise.</returns>
        private static bool CellFormatsAreEqual(CellFormat f, StyleInfo styleInfo, bool compareAlignment)
        {
            return
                f.BorderId != null && styleInfo.BorderId == f.BorderId
                && f.FillId != null && styleInfo.FillId == f.FillId
                && f.FontId != null && styleInfo.FontId == f.FontId
                && f.NumberFormatId != null && styleInfo.NumberFormatId == f.NumberFormatId
                && QuotePrefixesAreEqual(f.QuotePrefix, styleInfo.IncludeQuotePrefix)
                && (f.ApplyFill == null && styleInfo.Style.Fill == XLFillValue.Default ||
                    f.ApplyFill != null && f.ApplyFill == ApplyFill(styleInfo))
                && (f.ApplyBorder == null && styleInfo.Style.Border == XLBorderValue.Default ||
                    f.ApplyBorder != null && f.ApplyBorder == ApplyBorder(styleInfo))
                && (!compareAlignment || AlignmentsAreEqual(f.Alignment, styleInfo.Style.Alignment))
                && ProtectionsAreEqual(f.Protection, styleInfo.Style.Protection)
                ;
        }

        private static bool ProtectionsAreEqual(Protection protection, XLProtectionValue xlProtection)
        {
            var p = XLProtectionValue.Default.Key;
            if (protection != null)
            {
                if (protection.Locked != null)
                    p.Locked = protection.Locked.Value;
                if (protection.Hidden != null)
                    p.Hidden = protection.Hidden.Value;
            }
            return p.Equals(xlProtection.Key);
        }

        private static bool QuotePrefixesAreEqual(BooleanValue quotePrefix, Boolean includeQuotePrefix)
        {
            return OpenXmlHelper.GetBooleanValueAsBool(quotePrefix, false) == includeQuotePrefix;
        }

        private static bool AlignmentsAreEqual(Alignment alignment, XLAlignmentValue xlAlignment)
        {
            if (alignment != null)
            {
                var a = XLAlignmentValue.Default.Key;
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
                return a.Equals(xlAlignment.Key);
            }
            else
            {
                return XLStyle.Default.Value.Alignment.Equals(xlAlignment);
            }
        }

        private Dictionary<XLBorderValue, BorderInfo> ResolveBorders(WorkbookStylesPart workbookStylesPart,
            Dictionary<XLBorderValue, BorderInfo> sharedBorders)
        {
            if (workbookStylesPart.Stylesheet.Borders == null)
                workbookStylesPart.Stylesheet.Borders = new Borders();

            var allSharedBorders = new Dictionary<XLBorderValue, BorderInfo>();
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
            if (borderInfo.Border.DiagonalUp != XLBorderValue.Default.DiagonalUp || ignoreMod)
                border.DiagonalUp = borderInfo.Border.DiagonalUp;

            if (borderInfo.Border.DiagonalDown != XLBorderValue.Default.DiagonalDown || ignoreMod)
                border.DiagonalDown = borderInfo.Border.DiagonalDown;

            if (borderInfo.Border.LeftBorder != XLBorderValue.Default.LeftBorder || ignoreMod)
            {
                var leftBorder = new LeftBorder { Style = borderInfo.Border.LeftBorder.ToOpenXml() };
                if (borderInfo.Border.LeftBorderColor != XLBorderValue.Default.LeftBorderColor || ignoreMod)
                {
                    var leftBorderColor = new Color().FromClosedXMLColor<Color>(borderInfo.Border.LeftBorderColor);
                    leftBorder.AppendChild(leftBorderColor);
                }
                border.AppendChild(leftBorder);
            }

            if (borderInfo.Border.RightBorder != XLBorderValue.Default.RightBorder || ignoreMod)
            {
                var rightBorder = new RightBorder { Style = borderInfo.Border.RightBorder.ToOpenXml() };
                if (borderInfo.Border.RightBorderColor != XLBorderValue.Default.RightBorderColor || ignoreMod)
                {
                    var rightBorderColor = new Color().FromClosedXMLColor<Color>(borderInfo.Border.RightBorderColor);
                    rightBorder.AppendChild(rightBorderColor);
                }
                border.AppendChild(rightBorder);
            }

            if (borderInfo.Border.TopBorder != XLBorderValue.Default.TopBorder || ignoreMod)
            {
                var topBorder = new TopBorder { Style = borderInfo.Border.TopBorder.ToOpenXml() };
                if (borderInfo.Border.TopBorderColor != XLBorderValue.Default.TopBorderColor || ignoreMod)
                {
                    var topBorderColor = new Color().FromClosedXMLColor<Color>(borderInfo.Border.TopBorderColor);
                    topBorder.AppendChild(topBorderColor);
                }
                border.AppendChild(topBorder);
            }

            if (borderInfo.Border.BottomBorder != XLBorderValue.Default.BottomBorder || ignoreMod)
            {
                var bottomBorder = new BottomBorder { Style = borderInfo.Border.BottomBorder.ToOpenXml() };
                if (borderInfo.Border.BottomBorderColor != XLBorderValue.Default.BottomBorderColor || ignoreMod)
                {
                    var bottomBorderColor = new Color().FromClosedXMLColor<Color>(borderInfo.Border.BottomBorderColor);
                    bottomBorder.AppendChild(bottomBorderColor);
                }
                border.AppendChild(bottomBorder);
            }

            if (borderInfo.Border.DiagonalBorder != XLBorderValue.Default.DiagonalBorder || ignoreMod)
            {
                var DiagonalBorder = new DiagonalBorder { Style = borderInfo.Border.DiagonalBorder.ToOpenXml() };
                if (borderInfo.Border.DiagonalBorderColor != XLBorderValue.Default.DiagonalBorderColor || ignoreMod)
                    if (borderInfo.Border.DiagonalBorderColor != null)
                    {
                        var DiagonalBorderColor = new Color().FromClosedXMLColor<Color>(borderInfo.Border.DiagonalBorderColor);
                        DiagonalBorder.AppendChild(DiagonalBorderColor);
                    }
                border.AppendChild(DiagonalBorder);
            }

            return border;
        }

        private bool BordersAreEqual(Border b, XLBorderValue xlBorder)
        {
            var nb = XLBorderValue.Default.Key;
            if (b.DiagonalUp != null)
                nb.DiagonalUp = b.DiagonalUp.Value;

            if (b.DiagonalDown != null)
                nb.DiagonalDown = b.DiagonalDown.Value;

            if (b.DiagonalBorder != null)
            {
                if (b.DiagonalBorder.Style != null)
                    nb.DiagonalBorder = b.DiagonalBorder.Style.Value.ToClosedXml();
                if (b.DiagonalBorder.Color != null)
                    nb.DiagonalBorderColor = b.DiagonalBorder.Color.ToClosedXMLColor(_colorList).Key;
            }

            if (b.LeftBorder != null)
            {
                if (b.LeftBorder.Style != null)
                    nb.LeftBorder = b.LeftBorder.Style.Value.ToClosedXml();
                if (b.LeftBorder.Color != null)
                    nb.LeftBorderColor = b.LeftBorder.Color.ToClosedXMLColor(_colorList).Key;
            }

            if (b.RightBorder != null)
            {
                if (b.RightBorder.Style != null)
                    nb.RightBorder = b.RightBorder.Style.Value.ToClosedXml();
                if (b.RightBorder.Color != null)
                    nb.RightBorderColor = b.RightBorder.Color.ToClosedXMLColor(_colorList).Key;
            }

            if (b.TopBorder != null)
            {
                if (b.TopBorder.Style != null)
                    nb.TopBorder = b.TopBorder.Style.Value.ToClosedXml();
                if (b.TopBorder.Color != null)
                    nb.TopBorderColor = b.TopBorder.Color.ToClosedXMLColor(_colorList).Key;
            }

            if (b.BottomBorder != null)
            {
                if (b.BottomBorder.Style != null)
                    nb.BottomBorder = b.BottomBorder.Style.Value.ToClosedXml();
                if (b.BottomBorder.Color != null)
                    nb.BottomBorderColor = b.BottomBorder.Color.ToClosedXMLColor(_colorList).Key;
            }

            return nb.Equals(xlBorder.Key);
        }

        private Dictionary<XLFillValue, FillInfo> ResolveFills(WorkbookStylesPart workbookStylesPart,
            Dictionary<XLFillValue, FillInfo> sharedFills)
        {
            if (workbookStylesPart.Stylesheet.Fills == null)
                workbookStylesPart.Stylesheet.Fills = new Fills();

            ResolveFillWithPattern(workbookStylesPart.Stylesheet.Fills, PatternValues.None);
            ResolveFillWithPattern(workbookStylesPart.Stylesheet.Fills, PatternValues.Gray125);

            var allSharedFills = new Dictionary<XLFillValue, FillInfo>();
            foreach (var fillInfo in sharedFills.Values)
            {
                var fillId = 0;
                var foundOne = false;
                foreach (Fill f in workbookStylesPart.Stylesheet.Fills)
                {
                    if (FillsAreEqual(f, fillInfo.Fill, fromDifferentialFormat: false))
                    {
                        foundOne = true;
                        break;
                    }
                    fillId++;
                }
                if (!foundOne)
                {
                    var fill = GetNewFill(fillInfo, differentialFillFormat: false);
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
                f.PatternFill == null
                || (f.PatternFill.PatternType == patternValues
                    && f.PatternFill.ForegroundColor == null
                    && f.PatternFill.BackgroundColor == null
                )))
                return;

            var fill1 = new Fill();
            var patternFill1 = new PatternFill { PatternType = patternValues };
            fill1.AppendChild(patternFill1);
            fills.AppendChild(fill1);
        }

        private static Fill GetNewFill(FillInfo fillInfo, Boolean differentialFillFormat, Boolean ignoreMod = true)
        {
            var fill = new Fill();

            var patternFill = new PatternFill();

            patternFill.PatternType = fillInfo.Fill.PatternType.ToOpenXml();

            BackgroundColor backgroundColor;
            ForegroundColor foregroundColor;

            switch (fillInfo.Fill.PatternType)
            {
                case XLFillPatternValues.None:
                    break;

                case XLFillPatternValues.Solid:

                    if (differentialFillFormat)
                    {
                        patternFill.AppendChild(new ForegroundColor { Auto = true });
                        backgroundColor = new BackgroundColor().FromClosedXMLColor<BackgroundColor>(fillInfo.Fill.BackgroundColor, true);
                        if (backgroundColor.HasAttributes)
                            patternFill.AppendChild(backgroundColor);
                    }
                    else
                    {
                        // ClosedXML Background color to be populated into OpenXML fgColor
                        foregroundColor = new ForegroundColor().FromClosedXMLColor<ForegroundColor>(fillInfo.Fill.BackgroundColor);
                        if (foregroundColor.HasAttributes)
                            patternFill.AppendChild(foregroundColor);
                    }
                    break;

                default:

                    foregroundColor = new ForegroundColor().FromClosedXMLColor<ForegroundColor>(fillInfo.Fill.PatternColor);
                    if (foregroundColor.HasAttributes)
                        patternFill.AppendChild(foregroundColor);

                    backgroundColor = new BackgroundColor().FromClosedXMLColor<BackgroundColor>(fillInfo.Fill.BackgroundColor);
                    if (backgroundColor.HasAttributes)
                        patternFill.AppendChild(backgroundColor);

                    break;
            }

            if (patternFill.HasChildren)
                fill.AppendChild(patternFill);

            return fill;
        }

        private bool FillsAreEqual(Fill f, XLFillValue xlFill, Boolean fromDifferentialFormat)
        {
            var nF = new XLFill(null);

            LoadFill(f, nF, fromDifferentialFormat);

            return nF.Key.Equals(xlFill.Key);
        }

        private void ResolveFonts(WorkbookStylesPart workbookStylesPart, SaveContext context)
        {
            if (workbookStylesPart.Stylesheet.Fonts == null)
                workbookStylesPart.Stylesheet.Fonts = new Fonts();

            var newFonts = new Dictionary<XLFontValue, FontInfo>();
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
            var bold = (fontInfo.Font.Bold != XLFontValue.Default.Bold || ignoreMod) && fontInfo.Font.Bold ? new Bold() : null;
            var italic = (fontInfo.Font.Italic != XLFontValue.Default.Italic || ignoreMod) && fontInfo.Font.Italic ? new Italic() : null;
            var underline = (fontInfo.Font.Underline != XLFontValue.Default.Underline || ignoreMod) &&
                            fontInfo.Font.Underline != XLFontUnderlineValues.None
                ? new Underline { Val = fontInfo.Font.Underline.ToOpenXml() }
                : null;
            var strike = (fontInfo.Font.Strikethrough != XLFontValue.Default.Strikethrough || ignoreMod) && fontInfo.Font.Strikethrough
                ? new Strike()
                : null;
            var verticalAlignment = fontInfo.Font.VerticalAlignment != XLFontValue.Default.VerticalAlignment || ignoreMod
                ? new VerticalTextAlignment { Val = fontInfo.Font.VerticalAlignment.ToOpenXml() }
                : null;

            var shadow = (fontInfo.Font.Shadow != XLFontValue.Default.Shadow || ignoreMod) && fontInfo.Font.Shadow ? new Shadow() : null;
            var fontSize = fontInfo.Font.FontSize != XLFontValue.Default.FontSize || ignoreMod
                ? new FontSize { Val = fontInfo.Font.FontSize }
                : null;
            var color = fontInfo.Font.FontColor != XLFontValue.Default.FontColor || ignoreMod ? new Color().FromClosedXMLColor<Color>(fontInfo.Font.FontColor) : null;

            var fontName = fontInfo.Font.FontName != XLFontValue.Default.FontName || ignoreMod
                ? new FontName { Val = fontInfo.Font.FontName }
                : null;
            var fontFamilyNumbering = fontInfo.Font.FontFamilyNumbering != XLFontValue.Default.FontFamilyNumbering || ignoreMod
                ? new FontFamilyNumbering { Val = (Int32)fontInfo.Font.FontFamilyNumbering }
                : null;

            var fontCharSet = (fontInfo.Font.FontCharSet != XLFontValue.Default.FontCharSet || ignoreMod) && fontInfo.Font.FontCharSet != XLFontCharSet.Default
                ? new FontCharSet { Val = (Int32)fontInfo.Font.FontCharSet }
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
            if (fontCharSet != null)
                font.AppendChild(fontCharSet);

            return font;
        }

        private bool FontsAreEqual(Font f, XLFontValue xlFont)
        {
            var nf = XLFontValue.Default.Key;
            nf.Bold = f.Bold != null;
            nf.Italic = f.Italic != null;

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
            if (f.Color != null)
                nf.FontColor = f.Color.ToClosedXMLColor(_colorList).Key;
            if (f.FontName != null)
                nf.FontName = f.FontName.Val;
            if (f.FontFamilyNumbering != null)
                nf.FontFamilyNumbering = (XLFontFamilyNumberingValues)f.FontFamilyNumbering.Val.Value;

            return nf.Equals(xlFont.Key);
        }

        private static Dictionary<XLNumberFormatValue, NumberFormatInfo> ResolveNumberFormats(
            WorkbookStylesPart workbookStylesPart,
            Dictionary<XLNumberFormatValue, NumberFormatInfo> sharedNumberFormats,
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

            var allSharedNumberFormats = new Dictionary<XLNumberFormatValue, NumberFormatInfo>();
            foreach (var numberFormatInfo in sharedNumberFormats.Values.Where(nf => nf.NumberFormatId != defaultFormatId))
            {
                var numberingFormatId = XLConstants.NumberOfBuiltInStyles; // 0-based
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

        private static bool NumberFormatsAreEqual(NumberingFormat nf, XLNumberFormatValue xlNumberFormat)
        {
            if (nf.FormatCode != null && !String.IsNullOrWhiteSpace(nf.FormatCode.Value))
                return string.Equals(xlNumberFormat?.Format, nf.FormatCode.Value);
            else if (nf.NumberFormatId != null)
                return xlNumberFormat?.NumberFormatId == (Int32)nf.NumberFormatId.Value;
            return false;
        }

        #endregion GenerateWorkbookStylesPartContent
    }
}
