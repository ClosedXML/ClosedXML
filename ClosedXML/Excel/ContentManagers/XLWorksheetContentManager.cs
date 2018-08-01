using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

namespace ClosedXML.Excel.ContentManagers
{
    internal enum XLWorksheetContents
    {
        SheetProperties = 1,
        SheetDimension = 2,
        SheetViews = 3,
        SheetFormatProperties = 4,
        Columns = 5,
        SheetData = 6,
        SheetCalculationProperties = 7,
        SheetProtection = 8,
        ProtectedRanges = 9,
        Scenarios = 10,
        AutoFilter = 11,
        SortState = 12,
        DataConsolidate = 13,
        CustomSheetViews = 14,
        MergeCells = 15,
        PhoneticProperties = 16,
        ConditionalFormatting = 17,
        DataValidations = 18,
        Hyperlinks = 19,
        PrintOptions = 20,
        PageMargins = 21,
        PageSetup = 22,
        HeaderFooter = 23,
        RowBreaks = 24,
        ColumnBreaks = 25,
        CustomProperties = 26,
        CellWatches = 27,
        IgnoredErrors = 28,
        SmartTags = 29,
        Drawing = 30,
        LegacyDrawing = 31,
        LegacyDrawingHeaderFooter = 32,
        DrawingHeaderFooter = 33,
        Picture = 34,
        OleObjects = 35,
        Controls = 36,
        AlternateContent = 37,
        WebPublishItems = 38,
        TableParts = 39,
        WorksheetExtensionList = 40
    }

    internal class XLWorksheetContentManager : XLBaseContentManager<XLWorksheetContents>
    {
        public XLWorksheetContentManager(Worksheet opWorksheet)
        {
            contents.Add(XLWorksheetContents.SheetProperties, opWorksheet.Elements<SheetProperties>().LastOrDefault());
            contents.Add(XLWorksheetContents.SheetDimension, opWorksheet.Elements<SheetDimension>().LastOrDefault());
            contents.Add(XLWorksheetContents.SheetViews, opWorksheet.Elements<SheetViews>().LastOrDefault());
            contents.Add(XLWorksheetContents.SheetFormatProperties, opWorksheet.Elements<SheetFormatProperties>().LastOrDefault());
            contents.Add(XLWorksheetContents.Columns, opWorksheet.Elements<Columns>().LastOrDefault());
            contents.Add(XLWorksheetContents.SheetData, opWorksheet.Elements<SheetData>().LastOrDefault());
            contents.Add(XLWorksheetContents.SheetCalculationProperties, opWorksheet.Elements<SheetCalculationProperties>().LastOrDefault());
            contents.Add(XLWorksheetContents.SheetProtection, opWorksheet.Elements<SheetProtection>().LastOrDefault());
            contents.Add(XLWorksheetContents.ProtectedRanges, opWorksheet.Elements<ProtectedRanges>().LastOrDefault());
            contents.Add(XLWorksheetContents.Scenarios, opWorksheet.Elements<Scenarios>().LastOrDefault());
            contents.Add(XLWorksheetContents.AutoFilter, opWorksheet.Elements<AutoFilter>().LastOrDefault());
            contents.Add(XLWorksheetContents.SortState, opWorksheet.Elements<SortState>().LastOrDefault());
            contents.Add(XLWorksheetContents.DataConsolidate, opWorksheet.Elements<DataConsolidate>().LastOrDefault());
            contents.Add(XLWorksheetContents.CustomSheetViews, opWorksheet.Elements<CustomSheetViews>().LastOrDefault());
            contents.Add(XLWorksheetContents.MergeCells, opWorksheet.Elements<MergeCells>().LastOrDefault());
            contents.Add(XLWorksheetContents.PhoneticProperties, opWorksheet.Elements<PhoneticProperties>().LastOrDefault());
            contents.Add(XLWorksheetContents.ConditionalFormatting, opWorksheet.Elements<ConditionalFormatting>().LastOrDefault());
            contents.Add(XLWorksheetContents.DataValidations, opWorksheet.Elements<DataValidations>().LastOrDefault());
            contents.Add(XLWorksheetContents.Hyperlinks, opWorksheet.Elements<Hyperlinks>().LastOrDefault());
            contents.Add(XLWorksheetContents.PrintOptions, opWorksheet.Elements<PrintOptions>().LastOrDefault());
            contents.Add(XLWorksheetContents.PageMargins, opWorksheet.Elements<PageMargins>().LastOrDefault());
            contents.Add(XLWorksheetContents.PageSetup, opWorksheet.Elements<PageSetup>().LastOrDefault());
            contents.Add(XLWorksheetContents.HeaderFooter, opWorksheet.Elements<HeaderFooter>().LastOrDefault());
            contents.Add(XLWorksheetContents.RowBreaks, opWorksheet.Elements<RowBreaks>().LastOrDefault());
            contents.Add(XLWorksheetContents.ColumnBreaks, opWorksheet.Elements<ColumnBreaks>().LastOrDefault());
            contents.Add(XLWorksheetContents.CustomProperties, opWorksheet.Elements<CustomProperties>().LastOrDefault());
            contents.Add(XLWorksheetContents.CellWatches, opWorksheet.Elements<CellWatches>().LastOrDefault());
            contents.Add(XLWorksheetContents.IgnoredErrors, opWorksheet.Elements<IgnoredErrors>().LastOrDefault());
            //contents.Add(XLWSContents.SmartTags, opWorksheet.Elements<SmartTags>().LastOrDefault());
            contents.Add(XLWorksheetContents.Drawing, opWorksheet.Elements<Drawing>().LastOrDefault());
            contents.Add(XLWorksheetContents.LegacyDrawing, opWorksheet.Elements<LegacyDrawing>().LastOrDefault());
            contents.Add(XLWorksheetContents.LegacyDrawingHeaderFooter, opWorksheet.Elements<LegacyDrawingHeaderFooter>().LastOrDefault());
            contents.Add(XLWorksheetContents.DrawingHeaderFooter, opWorksheet.Elements<DrawingHeaderFooter>().LastOrDefault());
            contents.Add(XLWorksheetContents.Picture, opWorksheet.Elements<Picture>().LastOrDefault());
            contents.Add(XLWorksheetContents.OleObjects, opWorksheet.Elements<OleObjects>().LastOrDefault());
            contents.Add(XLWorksheetContents.Controls, opWorksheet.Elements<Controls>().LastOrDefault());
            contents.Add(XLWorksheetContents.AlternateContent, opWorksheet.Elements<AlternateContent>().LastOrDefault());
            contents.Add(XLWorksheetContents.WebPublishItems, opWorksheet.Elements<WebPublishItems>().LastOrDefault());
            contents.Add(XLWorksheetContents.TableParts, opWorksheet.Elements<TableParts>().LastOrDefault());
            contents.Add(XLWorksheetContents.WorksheetExtensionList, opWorksheet.Elements<WorksheetExtensionList>().LastOrDefault());
        }
    }
}
