using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Excel
{
    internal class XLWSContentManager
    {
        public enum XLWSContents
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
        private Dictionary<XLWSContents, OpenXmlElement> contents = new Dictionary<XLWSContents, OpenXmlElement>();

        public XLWSContentManager(Worksheet opWorksheet)
        {
            contents.Add(XLWSContents.SheetProperties, opWorksheet.Elements<SheetProperties>().FirstOrDefault());
            contents.Add(XLWSContents.SheetDimension, opWorksheet.Elements<SheetDimension>().FirstOrDefault());
            contents.Add(XLWSContents.SheetViews, opWorksheet.Elements<SheetViews>().FirstOrDefault());
            contents.Add(XLWSContents.SheetFormatProperties, opWorksheet.Elements<SheetFormatProperties>().FirstOrDefault());
            contents.Add(XLWSContents.Columns, opWorksheet.Elements<Columns>().FirstOrDefault());
            contents.Add(XLWSContents.SheetData, opWorksheet.Elements<SheetData>().FirstOrDefault());
            contents.Add(XLWSContents.SheetCalculationProperties, opWorksheet.Elements<SheetCalculationProperties>().FirstOrDefault());
            contents.Add(XLWSContents.SheetProtection, opWorksheet.Elements<SheetProtection>().FirstOrDefault());
            contents.Add(XLWSContents.ProtectedRanges, opWorksheet.Elements<ProtectedRanges>().FirstOrDefault());
            contents.Add(XLWSContents.Scenarios, opWorksheet.Elements<Scenarios>().FirstOrDefault());
            contents.Add(XLWSContents.AutoFilter, opWorksheet.Elements<AutoFilter>().FirstOrDefault());
            contents.Add(XLWSContents.SortState, opWorksheet.Elements<SortState>().FirstOrDefault());
            contents.Add(XLWSContents.DataConsolidate, opWorksheet.Elements<DataConsolidate>().FirstOrDefault());
            contents.Add(XLWSContents.CustomSheetViews, opWorksheet.Elements<CustomSheetViews>().FirstOrDefault());
            contents.Add(XLWSContents.MergeCells, opWorksheet.Elements<MergeCells>().FirstOrDefault());
            contents.Add(XLWSContents.PhoneticProperties, opWorksheet.Elements<PhoneticProperties>().FirstOrDefault());
            contents.Add(XLWSContents.ConditionalFormatting, opWorksheet.Elements<ConditionalFormatting>().FirstOrDefault());
            contents.Add(XLWSContents.DataValidations, opWorksheet.Elements<DataValidations>().FirstOrDefault());
            contents.Add(XLWSContents.Hyperlinks, opWorksheet.Elements<Hyperlinks>().FirstOrDefault());
            contents.Add(XLWSContents.PrintOptions, opWorksheet.Elements<PrintOptions>().FirstOrDefault());
            contents.Add(XLWSContents.PageMargins, opWorksheet.Elements<PageMargins>().FirstOrDefault());
            contents.Add(XLWSContents.PageSetup, opWorksheet.Elements<PageSetup>().FirstOrDefault());
            contents.Add(XLWSContents.HeaderFooter, opWorksheet.Elements<HeaderFooter>().FirstOrDefault());
            contents.Add(XLWSContents.RowBreaks, opWorksheet.Elements<RowBreaks>().FirstOrDefault());
            contents.Add(XLWSContents.ColumnBreaks, opWorksheet.Elements<ColumnBreaks>().FirstOrDefault());
            contents.Add(XLWSContents.CustomProperties, opWorksheet.Elements<CustomProperties>().FirstOrDefault());
            contents.Add(XLWSContents.CellWatches, opWorksheet.Elements<CellWatches>().FirstOrDefault());
            contents.Add(XLWSContents.IgnoredErrors, opWorksheet.Elements<IgnoredErrors>().FirstOrDefault());
            contents.Add(XLWSContents.SmartTags, opWorksheet.Elements<SmartTags>().FirstOrDefault());
            contents.Add(XLWSContents.Drawing, opWorksheet.Elements<Drawing>().FirstOrDefault());
            contents.Add(XLWSContents.LegacyDrawing, opWorksheet.Elements<LegacyDrawing>().FirstOrDefault());
            contents.Add(XLWSContents.LegacyDrawingHeaderFooter, opWorksheet.Elements<LegacyDrawingHeaderFooter>().FirstOrDefault());
            contents.Add(XLWSContents.DrawingHeaderFooter, opWorksheet.Elements<DrawingHeaderFooter>().FirstOrDefault());
            contents.Add(XLWSContents.Picture, opWorksheet.Elements<Picture>().FirstOrDefault());
            contents.Add(XLWSContents.OleObjects, opWorksheet.Elements<OleObjects>().FirstOrDefault());
            contents.Add(XLWSContents.Controls, opWorksheet.Elements<Controls>().FirstOrDefault());
            contents.Add(XLWSContents.AlternateContent, opWorksheet.Elements<AlternateContent>().FirstOrDefault());
            contents.Add(XLWSContents.WebPublishItems, opWorksheet.Elements<WebPublishItems>().FirstOrDefault());
            contents.Add(XLWSContents.TableParts, opWorksheet.Elements<TableParts>().FirstOrDefault());
            contents.Add(XLWSContents.WorksheetExtensionList, opWorksheet.Elements<WorksheetExtensionList>().FirstOrDefault());
        }

        public void SetElement(XLWSContents content, OpenXmlElement element)
        {
            contents[content] = element;
        }

        public OpenXmlElement GetPreviousElementFor(XLWSContents content)
        {
            var max = contents.Where(kp => (Int32)kp.Key < (Int32)content && kp.Value != null).Max(kp => kp.Key);
            return contents[max];
        }
    }
}
