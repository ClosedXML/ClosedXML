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
using BackgroundColor = DocumentFormat.OpenXml.Spreadsheet.BackgroundColor;
using Bold = DocumentFormat.OpenXml.Spreadsheet.Bold;
using Border = DocumentFormat.OpenXml.Spreadsheet.Border;
using BottomBorder = DocumentFormat.OpenXml.Spreadsheet.BottomBorder;
using Color = DocumentFormat.OpenXml.Spreadsheet.Color;
using Fill = DocumentFormat.OpenXml.Spreadsheet.Fill;
using Font = DocumentFormat.OpenXml.Spreadsheet.Font;
using FontCharSet = DocumentFormat.OpenXml.Spreadsheet.FontCharSet;
using Fonts = DocumentFormat.OpenXml.Spreadsheet.Fonts;
using FontSize = DocumentFormat.OpenXml.Spreadsheet.FontSize;
using ForegroundColor = DocumentFormat.OpenXml.Spreadsheet.ForegroundColor;
using Italic = DocumentFormat.OpenXml.Spreadsheet.Italic;
using LeftBorder = DocumentFormat.OpenXml.Spreadsheet.LeftBorder;
using NumberingFormat = DocumentFormat.OpenXml.Spreadsheet.NumberingFormat;
using Path = System.IO.Path;
using PatternFill = DocumentFormat.OpenXml.Spreadsheet.PatternFill;
using Properties = DocumentFormat.OpenXml.ExtendedProperties.Properties;
using RightBorder = DocumentFormat.OpenXml.Spreadsheet.RightBorder;
using Shadow = DocumentFormat.OpenXml.Spreadsheet.Shadow;
using Strike = DocumentFormat.OpenXml.Spreadsheet.Strike;
using TopBorder = DocumentFormat.OpenXml.Spreadsheet.TopBorder;
using Underline = DocumentFormat.OpenXml.Spreadsheet.Underline;
using VerticalTextAlignment = DocumentFormat.OpenXml.Spreadsheet.VerticalTextAlignment;
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

            ExtendedFilePropertiesPartWriter.GenerateContent(extendedFilePropertiesPart, this);

            GenerateWorkbookPartContent(workbookPart, options, context);

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
                    hasAnyVmlElements = VmlDrawingPartWriter.GenerateContent(vmlDrawingPart, worksheet);
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

    }
}
