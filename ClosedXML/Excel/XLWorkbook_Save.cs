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
