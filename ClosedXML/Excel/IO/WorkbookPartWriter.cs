#nullable disable

using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Utils;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel.IO
{
    internal class WorkbookPartWriter
    {
        internal static void GenerateContent(WorkbookPart workbookPart, XLWorkbook xlWorkbook, SaveOptions options, XLWorkbook.SaveContext context)
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

            workbook.WorkbookProperties.Date1904 = OpenXmlHelper.GetBooleanValue(xlWorkbook.Use1904DateSystem, false);

            if (options.FilterPrivacy.HasValue)
                workbook.WorkbookProperties.FilterPrivacy = OpenXmlHelper.GetBooleanValue(options.FilterPrivacy.Value, false);

            #endregion WorkbookProperties

            #region FileSharing

            if (workbook.FileSharing == null)
                workbook.FileSharing = new FileSharing();

            workbook.FileSharing.ReadOnlyRecommended = OpenXmlHelper.GetBooleanValue(xlWorkbook.FileSharing.ReadOnlyRecommended, false);
            workbook.FileSharing.UserName = String.IsNullOrWhiteSpace(xlWorkbook.FileSharing.UserName) ? null : StringValue.FromString(xlWorkbook.FileSharing.UserName);

            if (!workbook.FileSharing.HasChildren && !workbook.FileSharing.HasAttributes)
                workbook.FileSharing = null;

            #endregion FileSharing

            #region WorkbookProtection

            if (xlWorkbook.Protection.IsProtected)
            {
                if (workbook.WorkbookProtection == null)
                    workbook.WorkbookProtection = new WorkbookProtection();

                var workbookProtection = workbook.WorkbookProtection;

                var protection = xlWorkbook.Protection;

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

            var worksheets = xlWorkbook.WorksheetsInternal;
            workbook.Sheets.Elements<Sheet>().Where(s => worksheets.Deleted.Contains(s.Id)).ToList().ForEach(
                s => s.Remove());

            foreach (var sheet in workbook.Sheets.Elements<Sheet>())
            {
                var sheetId = (Int32)sheet.SheetId.Value;

                if (xlWorkbook.WorksheetsInternal.All<XLWorksheet>(w => w.SheetId != sheetId)) continue;

                var wks = xlWorkbook.WorksheetsInternal.Single<XLWorksheet>(w => w.SheetId == sheetId);
                wks.RelId = sheet.Id;
                sheet.Name = wks.Name;
            }

            foreach (var xlSheet in xlWorkbook.WorksheetsInternal.OrderBy<XLWorksheet, int>(w => w.Position))
            {
                string rId;
                if (String.IsNullOrWhiteSpace(xlSheet.RelId))
                {
                    // Sheet isn't from loaded file and hasn't been saved yet.
                    rId = xlSheet.RelId = context.RelIdGenerator.GetNext(XLWorkbook.RelType.Workbook);
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
                                join worksheet in ((IEnumerable<XLWorksheet>)xlWorkbook.WorksheetsInternal) on sheet.Id.Value
                                    equals worksheet.RelId
                                orderby worksheet.Position
                                select sheet;

            UInt32 firstSheetVisible = 0;
            var activeTab =
                (from us in xlWorkbook.UnsupportedSheets where us.IsActive select (UInt32)us.Position - 1).FirstOrDefault();
            var foundVisible = false;

            var totalSheets = sheetElements.Count() + xlWorkbook.UnsupportedSheets.Count;
            for (var p = 1; p <= totalSheets; p++)
            {
                if (xlWorkbook.UnsupportedSheets.All(us => us.Position != p))
                {
                    var sheet = sheetElements.ElementAt(p - xlWorkbook.UnsupportedSheets.Count(us => us.Position <= p) - 1);
                    workbook.Sheets.RemoveChild(sheet);
                    workbook.Sheets.AppendChild(sheet);
                    var xlSheet = xlWorkbook.Worksheet(sheet.Name);
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
                    var sheetId = xlWorkbook.UnsupportedSheets.First(us => us.Position == p).SheetId;
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
            foreach (var worksheet in xlWorkbook.WorksheetsInternal)
            {
                var wsSheetId = worksheet.SheetId;
                UInt32 sheetId = 0;
                foreach (var s in workbook.Sheets.Elements<Sheet>().TakeWhile(s => s.SheetId != wsSheetId))
                {
                    sheetId++;
                }

                if (!string.IsNullOrEmpty(worksheet.PageSetup.PrintAreas.PrintArea))
                {
                    var definedName = new DefinedName
                    {
                        Name = "_xlnm.Print_Area",
                        LocalSheetId = sheetId,
                        Text = worksheet.PageSetup.PrintAreas.PrintArea
                    };

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

                foreach (var xlDefinedName in worksheet.DefinedNames.Where<XLDefinedName>(n => n.Name != "_xlnm._FilterDatabase"))
                {
                    var definedName = new DefinedName
                    {
                        Name = xlDefinedName.Name,
                        LocalSheetId = sheetId,
                        Text = xlDefinedName.ToString()
                    };

                    if (!xlDefinedName.Visible)
                        definedName.Hidden = BooleanValue.FromBoolean(true);

                    if (!String.IsNullOrWhiteSpace(xlDefinedName.Comment))
                        definedName.Comment = xlDefinedName.Comment;

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

            foreach (var xlDefinedName in xlWorkbook.DefinedNamesInternal)
            {
                var definedName = new DefinedName
                {
                    Name = xlDefinedName.Name,
                    Text = xlDefinedName.RefersTo
                };

                if (!xlDefinedName.Visible)
                    definedName.Hidden = BooleanValue.FromBoolean(true);

                if (!String.IsNullOrWhiteSpace(xlDefinedName.Comment))
                    definedName.Comment = xlDefinedName.Comment;
                definedNames.AppendChild(definedName);
            }

            workbook.DefinedNames = definedNames;

            if (workbook.CalculationProperties == null)
                workbook.CalculationProperties = new CalculationProperties { CalculationId = 125725U };

            if (xlWorkbook.CalculateMode == XLCalculateMode.Default)
                workbook.CalculationProperties.CalculationMode = null;
            else
                workbook.CalculationProperties.CalculationMode = xlWorkbook.CalculateMode.ToOpenXml();

            if (xlWorkbook.ReferenceStyle == XLReferenceStyle.Default)
                workbook.CalculationProperties.ReferenceMode = null;
            else
                workbook.CalculationProperties.ReferenceMode = xlWorkbook.ReferenceStyle.ToOpenXml();

            if (xlWorkbook.CalculationOnSave) workbook.CalculationProperties.CalculationOnSave = xlWorkbook.CalculationOnSave;
            if (xlWorkbook.ForceFullCalculation) workbook.CalculationProperties.ForceFullCalculation = xlWorkbook.ForceFullCalculation;
            if (xlWorkbook.FullCalculationOnLoad) workbook.CalculationProperties.FullCalculationOnLoad = xlWorkbook.FullCalculationOnLoad;
            if (xlWorkbook.FullPrecision) workbook.CalculationProperties.FullPrecision = xlWorkbook.FullPrecision;
        }
    }
}
