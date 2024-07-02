#nullable disable

using ClosedXML.Extensions;
using ClosedXML.Utils;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using Path = System.IO.Path;
using ClosedXML.Excel.IO;
using Boolean = System.Boolean;
using System.Diagnostics;

namespace ClosedXML.Excel
{
    public partial class XLWorkbook
    {
        private Boolean Validate(SpreadsheetDocument package)
        {
            var backupCulture = Thread.CurrentThread.CurrentCulture;

            IList<ValidationErrorInfo> errors;
            try
            {
                Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
                var validator = new OpenXmlValidator();
                errors = validator.Validate(package).ToArray();
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

            ExtendedFilePropertiesPartWriter.GenerateContent(extendedFilePropertiesPart, this);

            WorkbookPartWriter.GenerateContent(workbookPart, this, options, context);

            var sharedStringTablePart = workbookPart.SharedStringTablePart ??
                                        workbookPart.AddNewPart<SharedStringTablePart>(
                                            context.RelIdGenerator.GetNext(RelType.Workbook));

            SharedStringTableWriter.GenerateSharedStringTablePartContent(this, sharedStringTablePart, context);

            var workbookStylesPart = workbookPart.WorkbookStylesPart ??
                                     workbookPart.AddNewPart<WorkbookStylesPart>(
                                         context.RelIdGenerator.GetNext(RelType.Workbook));

            WorkbookStylesPartWriter.GenerateContent(workbookStylesPart, this, context);

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

                // VML part is the source of truth for shapes of comments, form controls and likely others.
                // Excel won't display any shape without VML. The drawing part is always present, but is likely
                // only different rendering of VML (more precisely the shapes behind VML).
                var vmlDrawingPart = worksheetPart.VmlDrawingParts.FirstOrDefault();
                var hasAnyVmlElements = DeleteExistingCommentsShapes(vmlDrawingPart);

                if (worksheetHasComments)
                {
                    // If sheet has comments, we must keep VML in legacy drawing part to display them
                    // as well as comments part for semantic reasons.
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
                    hasAnyVmlElements = VmlDrawingPartWriter.GenerateContent(vmlDrawingPart, worksheet);
                }
                else
                {
                    // There are no comments in the worksheet = the comment part is no longer needed,
                    // but VML part might contain other shapes, like form controls.
                    if (commentsPart is not null)
                        worksheetPart.DeletePart(commentsPart);
                }

                if (!hasAnyVmlElements && vmlDrawingPart is not null)
                {
                    worksheet.LegacyDrawingId = null;
                    worksheetPart.DeletePart(vmlDrawingPart);
                }

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

                if (worksheet.PivotTables.Any<XLPivotTable>())
                {
                    GeneratePivotTables(workbookPart, worksheetPart, worksheet, context);
                }
            }

            if (options.GenerateCalculationChain)
            {
                CalculationChainPartWriter.GenerateContent(workbookPart, this, context);
            }
            else
            {
                if (workbookPart.CalculationChainPart is not null)
                    workbookPart.DeletePart(workbookPart.CalculationChainPart);
            }

            if (workbookPart.ThemePart == null)
            {
                var themePart = workbookPart.AddNewPart<ThemePart>(context.RelIdGenerator.GetNext(RelType.Workbook));
                ThemePartWriter.GenerateContent(themePart, (XLTheme)Theme);
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

            var pivotTables = WorksheetsInternal.SelectMany<XLWorksheet, XLPivotTable>(ws => ws.PivotTables);

            var xlPivotCaches = pivotTables.Select(pt => pt.PivotCache).Distinct();
            foreach (var xlPivotCache in xlPivotCaches)
            {
                Debug.Assert(workbookPart.Workbook.PivotCaches is not null);
                Debug.Assert(!string.IsNullOrEmpty(xlPivotCache.WorkbookCacheRelId));

                var pivotTableCacheDefinitionPart = (PivotTableCacheDefinitionPart)workbookPart.GetPartById(xlPivotCache.WorkbookCacheRelId);

                PivotTableCacheDefinitionPartWriter.GenerateContent(pivotTableCacheDefinitionPart, xlPivotCache, context);

                var pivotTableCacheRecordsPart = pivotTableCacheDefinitionPart.GetPartsOfType<PivotTableCacheRecordsPart>().Any()
                    ? pivotTableCacheDefinitionPart.GetPartsOfType<PivotTableCacheRecordsPart>().Single()
                    : pivotTableCacheDefinitionPart.AddNewPart<PivotTableCacheRecordsPart>("rId1");

                PivotTableCacheRecordsPartWriter.WriteContent(pivotTableCacheRecordsPart, xlPivotCache);
            }
        }

        private static void GeneratePivotTables(
            WorkbookPart workbookPart,
            WorksheetPart worksheetPart,
            XLWorksheet xlWorksheet,
            SaveContext context)
        {
            foreach (var pt in xlWorksheet.PivotTables)
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
                    pivotTablePart = (PivotTablePart)worksheetPart.GetPartById(pt.RelId);

                var pivotSource = pt.PivotCache;
                var pivotTableCacheDefinitionPart = pivotTablePart.PivotTableCacheDefinitionPart;
                if (!workbookPart.GetPartById(pivotSource.WorkbookCacheRelId).Equals(pivotTableCacheDefinitionPart))
                {
                    pivotTablePart.DeletePart(pivotTableCacheDefinitionPart);
                    pivotTablePart.CreateRelationshipToPart(workbookPart.GetPartById(pivotSource.WorkbookCacheRelId), context.RelIdGenerator.GetNext(XLWorkbook.RelType.Workbook));
                }

                PivotTableDefinitionPartWriter2.WriteContent(pivotTablePart, pt, context);
            }
        }

    }
}
