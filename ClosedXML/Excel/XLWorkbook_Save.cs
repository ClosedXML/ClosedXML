#nullable disable

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
using System.Diagnostics;
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
using Field = DocumentFormat.OpenXml.Spreadsheet.Field;
using Fill = DocumentFormat.OpenXml.Spreadsheet.Fill;
using Font = DocumentFormat.OpenXml.Spreadsheet.Font;
using FontCharSet = DocumentFormat.OpenXml.Spreadsheet.FontCharSet;
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
using Outline = DocumentFormat.OpenXml.Drawing.Outline;
using Path = System.IO.Path;
using PatternFill = DocumentFormat.OpenXml.Spreadsheet.PatternFill;
using Properties = DocumentFormat.OpenXml.ExtendedProperties.Properties;
using RightBorder = DocumentFormat.OpenXml.Spreadsheet.RightBorder;
using Shadow = DocumentFormat.OpenXml.Spreadsheet.Shadow;
using Strike = DocumentFormat.OpenXml.Spreadsheet.Strike;
using TopBorder = DocumentFormat.OpenXml.Spreadsheet.TopBorder;
using Underline = DocumentFormat.OpenXml.Spreadsheet.Underline;
using VerticalTextAlignment = DocumentFormat.OpenXml.Spreadsheet.VerticalTextAlignment;
using Vml = DocumentFormat.OpenXml.Vml;
using ClosedXML.Excel.Cells;
using ClosedXML.Excel.IO;
using Boolean = System.Boolean;

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

            GenerateExtendedFilePropertiesPartContent(extendedFilePropertiesPart, this);

            GenerateWorkbookPartContent(workbookPart, options, context);

            var sharedStringTablePart = workbookPart.SharedStringTablePart ??
                                        workbookPart.AddNewPart<SharedStringTablePart>(
                                            context.RelIdGenerator.GetNext(RelType.Workbook));

            SharedStringTableWriter.GenerateSharedStringTablePartContent(this, sharedStringTablePart, context);

            var workbookStylesPart = workbookPart.WorkbookStylesPart ??
                                     workbookPart.AddNewPart<WorkbookStylesPart>(
                                         context.RelIdGenerator.GetNext(RelType.Workbook));

            GenerateWorkbookStylesPartContent(workbookStylesPart, context);

            var cacheRelIds = PivotCachesInternal
                  .Select<XLPivotCache, String>(ps => ps.WorkbookCacheRelId)
                  .Where(relId => !string.IsNullOrWhiteSpace(relId))
                  .Distinct();

            foreach (var relId in cacheRelIds)
            {
                if (workbookPart.GetPartById(relId) is PivotTableCacheDefinitionPart pivotTableCacheDefinitionPart)
                    pivotTableCacheDefinitionPart.PivotCacheDefinition.CacheFields.RemoveAllChildren();
            }

            var allPivotTables = WorksheetsInternal.SelectMany<XLWorksheet, IXLPivotTable>(ws => ws.PivotTables).ToList();

            // Phase 1 - Synchronize all pivot cache parts in the document, so each
            // source that will be saved has all required parts created and relationship
            // ids are set (in this case `Workbook.PivotCaches` relationship table).
            // Only sources that are used by a table are saved.
            SynchronizePivotTableParts(workbookPart, allPivotTables, context);

            // Phase 2 - All parts and relationships are set, fill in the parts.
            if (allPivotTables.Any())
            {
                GeneratePivotCaches(workbookPart, context);
            }

            foreach (var worksheet in WorksheetsInternal.Cast<XLWorksheet>().OrderBy(w => w.Position))
            {
                WorksheetPart worksheetPart;
                var wsRelId = worksheet.RelId;
                bool partIsEmpty;
                if (workbookPart.Parts.Any(p => p.RelationshipId == wsRelId))
                {
                    worksheetPart = (WorksheetPart)workbookPart.GetPartById(wsRelId);
                    partIsEmpty = false;
                }
                else
                {
                    worksheetPart = workbookPart.AddNewPart<WorksheetPart>(wsRelId);
                    partIsEmpty = true;
                }

                var worksheetHasComments = worksheet.Internals.CellsCollection.GetCells(c => c.HasComment).Any();

                var commentsPart = worksheetPart.WorksheetCommentsPart;
                var vmlDrawingPart = worksheetPart.VmlDrawingParts.FirstOrDefault();
                var hasAnyVmlElements = DeleteExistingCommentsShapes(vmlDrawingPart);

                if (worksheetHasComments)
                {
                    if (commentsPart == null)
                    {
                        commentsPart = worksheetPart.AddNewPart<WorksheetCommentsPart>(context.RelIdGenerator.GetNext(RelType.Workbook));
                    }

                    if (vmlDrawingPart == null)
                    {
                        if (String.IsNullOrWhiteSpace(worksheet.LegacyDrawingId))
                        {
                            worksheet.LegacyDrawingId = context.RelIdGenerator.GetNext(RelType.Workbook);
                        }

                        vmlDrawingPart = worksheetPart.AddNewPart<VmlDrawingPart>(worksheet.LegacyDrawingId);
                    }

                    CommentPartWriter.GenerateWorksheetCommentsPartContent(commentsPart, worksheet);
                    hasAnyVmlElements = GenerateVmlDrawingPartContent(vmlDrawingPart, worksheet);
                }
                else
                {
                    worksheet.LegacyDrawingId = null;
                    if (commentsPart is not null)
                        worksheetPart.DeletePart(commentsPart);
                }

                if (!hasAnyVmlElements && vmlDrawingPart != null)
                    worksheetPart.DeletePart(vmlDrawingPart);

                var xlTables = worksheet.Tables;

                // The way forward is to have 2-phase save, this is a start of that
                // concept for tables:
                //
                // Phase 1 - synchronize part existence with tables xlWorksheet, so each
                // table has a corresponding part and part that don't are deleted.
                // This phase doesn't modify the content, it only ensures that RelIds are set
                // corresponding parts exist and the parts that don't exist are removed
                TablePartWriter.SynchronizeTableParts(xlTables, worksheetPart, context);

                // Phase 2 - At this point, all pieces must have corresponding parts
                // The only way to link between parts is through RelIds that were already
                // set in phase 1. The phase 2 is all about content of individual parts.
                // Each part should have individual writer.
                TablePartWriter.GenerateTableParts(xlTables, worksheetPart, context);

                WorksheetPartWriter.GenerateWorksheetPartContent(partIsEmpty, worksheetPart, worksheet, options, context);

                if (worksheet.PivotTables.Any())
                {
                    GeneratePivotTables(workbookPart, worksheetPart, worksheet, context);
                }
            }

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

                CustomFilePropertiesPartWriter.GenerateContent(customFilePropertiesPart, this);
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

        private bool DeleteExistingCommentsShapes(VmlDrawingPart vmlDrawingPart)
        {
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

        private static void GenerateExtendedFilePropertiesPartContent(ExtendedFilePropertiesPart extendedFilePropertiesPart, XLWorkbook workbook)
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
                ((IEnumerable<XLWorksheet>)workbook.WorksheetsInternal).Select(w => new { w.Name, Order = w.Position }).ToList();
            var modifiedNamedRanges = GetModifiedNamedRanges(workbook);
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

            if (workbook.Properties.Manager != null)
            {
                if (!String.IsNullOrWhiteSpace(workbook.Properties.Manager))
                {
                    if (properties.Manager == null)
                        properties.Manager = new Manager();

                    properties.Manager.Text = workbook.Properties.Manager;
                }
                else
                    properties.Manager = null;
            }

            if (workbook.Properties.Company == null) return;

            if (!String.IsNullOrWhiteSpace(workbook.Properties.Company))
            {
                if (properties.Company == null)
                    properties.Company = new Company();

                properties.Company.Text = workbook.Properties.Company;
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

        private static List<string> GetModifiedNamedRanges(XLWorkbook workbook)
        {
            var namedRanges = new List<String>();
            foreach (var w in workbook.WorksheetsInternal)
            {
                var wName = w.Name;
                namedRanges.AddRange(w.NamedRanges.Select(n => wName + "!" + n.Name));
                namedRanges.Add(w.Name + "!Print_Area");
                namedRanges.Add(w.Name + "!Print_Titles");
            }
            namedRanges.AddRange(workbook.NamedRanges.Select(n => n.Name));
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

            foreach (var xlSheet in WorksheetsInternal.OrderBy<XLWorksheet, int>(w => w.Position))
            {
                string rId;
                if (String.IsNullOrWhiteSpace(xlSheet.RelId))
                {
                    // Sheet isn't from loaded file and hasn't been saved yet.
                    rId = xlSheet.RelId = context.RelIdGenerator.GetNext(RelType.Workbook);
                }
                else
                {
                    // Keep same r:id from previous file
                    rId = xlSheet.RelId;
                }

                if (workbook.Sheets.Cast<Sheet>().All(s => s.Id != rId))
                {
                    var newSheet = new Sheet
                    {
                        Name = xlSheet.Name,
                        Id = rId,
                        SheetId = xlSheet.SheetId
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
                var wsSheetId = worksheet.SheetId;
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
                    if (c.Formula.Type == FormulaType.DataTable)
                    {
                        // Do nothing, Excel doesn't generate calc chain for data table
                    }
                    else if (c.HasArrayFormula)
                    {
                        if (c.FormulaReference == null)
                            c.FormulaReference = c.AsRange().RangeAddress;

                        if (c.FormulaReference.FirstAddress.Equals(c.Address))
                        {
                            var cc = new CalculationCell
                            {
                                CellReference = c.Address.ToString(),
                                SheetId = (Int32)worksheet.SheetId
                            };

                            cc.Array = true;
                            calculationChain.AppendChild(cc);

                            foreach (var childCell in worksheet.Range(c.FormulaReference).Cells())
                            {
                                calculationChain.AppendChild(new CalculationCell
                                {
                                    CellReference = childCell.Address.ToString(),
                                    SheetId = (Int32)worksheet.SheetId,
                                });
                            }
                        }
                    }
                    else
                    {
                        calculationChain.AppendChild(new CalculationCell
                        {
                            CellReference = c.Address.ToString(),
                            SheetId = (Int32)worksheet.SheetId
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

        private static void SynchronizePivotTableParts(WorkbookPart workbookPart, IReadOnlyList<IXLPivotTable> allPivotTables, SaveContext context)
        {
            RemoveUnusedPivotCacheDefinitionParts(workbookPart, allPivotTables);
            AddUsedPivotCacheDefinitionParts(workbookPart, allPivotTables, context);

            // Ensure this in workbook.xml:
            //  <pivotCaches>
            //    <pivotCache cacheId="13" r:id="rId3"/>
            //  </pivotCaches>

            context.PivotSourceCacheId = 0;
            var xlUsedCaches = allPivotTables.Select(pt => pt.PivotCache).Distinct().Cast<XLPivotCache>().ToList();
            if (xlUsedCaches.Any())
            {
                // Recreate the workbook pivot cache references to remove previous gunk
                var pivotCaches = new PivotCaches();
                workbookPart.Workbook.PivotCaches = pivotCaches;

                foreach (var source in xlUsedCaches)
                {
                    var cacheId = context.PivotSourceCacheId++;
                    source.CacheId = cacheId;
                    var pivotCache = new PivotCache { CacheId = cacheId, Id = source.WorkbookCacheRelId };
                    pivotCaches.AppendChild(pivotCache);
                }
            }
            else
            {
                // Remove empty pivot cache part
                if (workbookPart.Workbook.PivotCaches is not null)
                {
                    workbookPart.Workbook.RemoveChild(workbookPart.Workbook.PivotCaches);
                }
            }

            // Remove pivot cache parts that are a part of the loaded document, but aren't used by a pivot table of the xlWorkbook
            // part of the first phase of saving
            static void RemoveUnusedPivotCacheDefinitionParts(WorkbookPart workbookPart, IReadOnlyList<IXLPivotTable> allPivotTables)
            {
                var workbookCacheRelIds = allPivotTables
                    .Select(pt => pt.PivotCache.CastTo<XLPivotCache>().WorkbookCacheRelId)
                    .Distinct()
                    .ToList();

                var orphanedParts = workbookPart
                    .GetPartsOfType<PivotTableCacheDefinitionPart>()
                    .Where(pcdp => !workbookCacheRelIds.Contains(workbookPart.GetIdOfPart(pcdp)))
                    .ToList();

                foreach (var orphanPart in orphanedParts)
                {
                    orphanPart.DeletePart(orphanPart.PivotTableCacheRecordsPart);
                    workbookPart.DeletePart(orphanPart);
                };

                // Remove deleted pivot cache parts
                if (workbookPart.Workbook.PivotCaches is not null)
                {
                    workbookPart.Workbook.PivotCaches.Elements<PivotCache>()
                        .Where(pc => !workbookPart.HasPartWithId(pc.Id))
                        .ToList()
                        .ForEach(pc => pc.Remove());
                }
            }

            static void AddUsedPivotCacheDefinitionParts(WorkbookPart workbookPart, IReadOnlyList<IXLPivotTable> allPivotTables, SaveContext context)
            {
                // Add ids and part for the caches to workbooks
                // We might get a XLPivotSource with an id of apart that isn't in the file (e.g. loaded from a file and saved to a different one).
                var newPivotSources = allPivotTables
                    .Select(pt => pt.PivotCache.CastTo<XLPivotCache>())
                    .Where(ps => string.IsNullOrEmpty(ps.WorkbookCacheRelId) || !workbookPart.HasPartWithId(ps.WorkbookCacheRelId))
                    .Distinct()
                    .ToList();

                foreach (var pivotSource in newPivotSources)
                {
                    var cacheRelId = context.RelIdGenerator.GetNext(RelType.Workbook);
                    pivotSource.WorkbookCacheRelId = cacheRelId;

                    workbookPart.AddNewPart<PivotTableCacheDefinitionPart>(pivotSource.WorkbookCacheRelId);
                }
            }
        }

        private void GeneratePivotCaches(WorkbookPart workbookPart, SaveContext context)
        {
            context.PivotSources.Clear();

            var pivotTables = WorksheetsInternal.Cast<XLWorksheet>().SelectMany(ws => ws.PivotTables);

            var pivotSources = pivotTables.Select(pt => pt.PivotCache).Distinct();
            foreach (var pivotSource in pivotSources.Cast<XLPivotCache>())
            {
                PivotTableCacheDefinitionPartWriter.GenerateContent(workbookPart, pivotSource, context);
            }
        }

        private static void GeneratePivotTables(
            WorkbookPart workbookPart,
            WorksheetPart worksheetPart,
            XLWorksheet xlWorksheet,
            SaveContext context)
        {
            foreach (var pt in xlWorksheet.PivotTables.Cast<XLPivotTable>())
            {
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

                PivotTablePartWriter.GeneratePivotTablePartContent(workbookPart, pivotTablePart, pt, context);
            }
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

                foreach (var c in worksheet.Internals.CellsCollection.GetCells())
                {
                    xlStyles.Add(c.StyleValue);
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

            var customNumberFormats = numberFormats
                .Where(nf => nf.NumberFormatId == -1)
                .ToHashSet();

            foreach (var pivotNumberFormat in pivotTableNumberFormats.Where(nf => nf.NumberFormatId == -1))
            {
                var numberFormatKey = new XLNumberFormatKey
                {
                    NumberFormatId = -1,
                    Format = pivotNumberFormat.Format
                };
                var numberFormat = XLNumberFormatValue.FromKey(ref numberFormatKey);

                customNumberFormats.Add(numberFormat);
            }

            var allSharedNumberFormats = ResolveNumberFormats(workbookStylesPart, customNumberFormats, defaultFormatId);
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
                    TextRotation = (UInt32)GetOpenXmlTextRotation(styleInfo.Style.Alignment),
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

            static int GetOpenXmlTextRotation(XLAlignmentValue alignment)
            {
                var textRotation = alignment.TextRotation;
                return textRotation >= 0
                    ? textRotation
                    : 90 - textRotation;
            }
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
                    a.TextRotation = OpenXmlHelper.GetClosedXmlTextRotation(alignment);
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

            var fills = workbookStylesPart.Stylesheet.Fills;

            // Pattern idx 0 and idx 1 are hardcoded to Excel with values None (0) and Gray125. Excel will ignore
            // values from the file. Every file has have these values inside to keep the first available idx at 2.
            ResolveFillWithPattern(fills, 0, PatternValues.None);
            ResolveFillWithPattern(fills, 1, PatternValues.Gray125);

            var allSharedFills = new Dictionary<XLFillValue, FillInfo>();
            foreach (var fillInfo in sharedFills.Values)
            {
                var fillId = 0;
                var foundOne = false;
                foreach (Fill f in fills)
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
                    fills.AppendChild(fill);
                }
                allSharedFills.Add(fillInfo.Fill, new FillInfo { Fill = fillInfo.Fill, FillId = (UInt32)fillId });
            }

            fills.Count = (UInt32)fills.Count();
            return allSharedFills;
        }

        private static void ResolveFillWithPattern(Fills fills, Int32 index, PatternValues patternValues)
        {
            var fill = (Fill)fills.ElementAtOrDefault(index);
            if (fill is null)
            {
                fills.InsertAt(new Fill { PatternFill = new PatternFill { PatternType = patternValues } }, index);
                return;
            }

            var fillHasExpectedValue =
                fill.PatternFill?.PatternType?.Value == patternValues &&
                fill.PatternFill.ForegroundColor is null &&
                fill.PatternFill.BackgroundColor is null;

            if (fillHasExpectedValue)
                return;

            fill.PatternFill = new PatternFill { PatternType = patternValues };
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

            var fontScheme = (fontInfo.Font.FontScheme != XLFontValue.Default.FontScheme || ignoreMod) && fontInfo.Font.FontScheme != XLFontScheme.None
                ? new DocumentFormat.OpenXml.Spreadsheet.FontScheme { Val = fontInfo.Font.FontScheme.ToOpenXmlEnum() }
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
            if (fontScheme != null)
                font.AppendChild(fontScheme);

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
            if (f.FontScheme?.Val != null)
                nf.FontScheme = f.FontScheme.Val.Value.ToClosedXml();

            return nf.Equals(xlFont.Key);
        }

        private static Dictionary<XLNumberFormatValue, NumberFormatInfo> ResolveNumberFormats(
            WorkbookStylesPart workbookStylesPart,
            HashSet<XLNumberFormatValue> customNumberFormats,
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
            var partNumberingFormats = workbookStylesPart.Stylesheet.NumberingFormats;

            // number format ids in the part can have holes in the sequence and first id can be greater than last built-in style id.
            // In some cases, there are also existing number formats with id below last built-in style id.
            var availableNumberFormatId = partNumberingFormats.Any()
                ? Math.Max(partNumberingFormats.Cast<NumberingFormat>().Max(nf => nf.NumberFormatId!.Value) + 1, XLConstants.NumberOfBuiltInStyles)
                : XLConstants.NumberOfBuiltInStyles; // 0-based

            // Merge custom formats used in the workbook that are not already present in the part to the part and assign ids
            foreach (var customNumberFormat in customNumberFormats.Where(nf => nf.NumberFormatId != defaultFormatId))
            {
                NumberingFormat partNumberFormat = null;
                foreach (var nf in workbookStylesPart.Stylesheet.NumberingFormats.Cast<NumberingFormat>())
                {
                    if (CustomNumberFormatsAreEqual(nf, customNumberFormat))
                    {
                        partNumberFormat = nf;
                        break;
                    }
                }
                if (partNumberFormat is null)
                {
                    partNumberFormat = new NumberingFormat
                    {
                        NumberFormatId = (UInt32)availableNumberFormatId++,
                        FormatCode = customNumberFormat.Format
                    };
                    workbookStylesPart.Stylesheet.NumberingFormats.AppendChild(partNumberFormat);
                }
                allSharedNumberFormats.Add(customNumberFormat,
                    new NumberFormatInfo
                    {
                        NumberFormat = customNumberFormat,
                        NumberFormatId = (Int32)partNumberFormat.NumberFormatId!.Value
                    });
            }
            workbookStylesPart.Stylesheet.NumberingFormats.Count =
                (UInt32)workbookStylesPart.Stylesheet.NumberingFormats.Count();
            return allSharedNumberFormats;
        }

        private static bool CustomNumberFormatsAreEqual(NumberingFormat nf, XLNumberFormatValue xlNumberFormat)
        {
            if (nf.FormatCode != null && !String.IsNullOrWhiteSpace(nf.FormatCode.Value))
                return string.Equals(xlNumberFormat?.Format, nf.FormatCode.Value);

            return false;
        }

        #endregion GenerateWorkbookStylesPartContent
    }
}
