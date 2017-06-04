using System;
using System.Collections.Generic;
using System.Linq;
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
            contents.Add(XLWSContents.SheetProperties, opWorksheet.Elements<SheetProperties>().LastOrDefault());
            contents.Add(XLWSContents.SheetDimension, opWorksheet.Elements<SheetDimension>().LastOrDefault());
            contents.Add(XLWSContents.SheetViews, opWorksheet.Elements<SheetViews>().LastOrDefault());
            contents.Add(XLWSContents.SheetFormatProperties, opWorksheet.Elements<SheetFormatProperties>().LastOrDefault());
            contents.Add(XLWSContents.Columns, opWorksheet.Elements<Columns>().LastOrDefault());
            contents.Add(XLWSContents.SheetData, opWorksheet.Elements<SheetData>().LastOrDefault());
            contents.Add(XLWSContents.SheetCalculationProperties, opWorksheet.Elements<SheetCalculationProperties>().LastOrDefault());
            contents.Add(XLWSContents.SheetProtection, opWorksheet.Elements<SheetProtection>().LastOrDefault());
            contents.Add(XLWSContents.ProtectedRanges, opWorksheet.Elements<ProtectedRanges>().LastOrDefault());
            contents.Add(XLWSContents.Scenarios, opWorksheet.Elements<Scenarios>().LastOrDefault());
            contents.Add(XLWSContents.AutoFilter, opWorksheet.Elements<AutoFilter>().LastOrDefault());
            contents.Add(XLWSContents.SortState, opWorksheet.Elements<SortState>().LastOrDefault());
            contents.Add(XLWSContents.DataConsolidate, opWorksheet.Elements<DataConsolidate>().LastOrDefault());
            contents.Add(XLWSContents.CustomSheetViews, opWorksheet.Elements<CustomSheetViews>().LastOrDefault());
            contents.Add(XLWSContents.MergeCells, opWorksheet.Elements<MergeCells>().LastOrDefault());
            contents.Add(XLWSContents.PhoneticProperties, opWorksheet.Elements<PhoneticProperties>().LastOrDefault());
            contents.Add(XLWSContents.ConditionalFormatting, opWorksheet.Elements<ConditionalFormatting>().LastOrDefault());
            contents.Add(XLWSContents.DataValidations, opWorksheet.Elements<DataValidations>().LastOrDefault());
            contents.Add(XLWSContents.Hyperlinks, opWorksheet.Elements<Hyperlinks>().LastOrDefault());
            contents.Add(XLWSContents.PrintOptions, opWorksheet.Elements<PrintOptions>().LastOrDefault());
            contents.Add(XLWSContents.PageMargins, opWorksheet.Elements<PageMargins>().LastOrDefault());
            contents.Add(XLWSContents.PageSetup, opWorksheet.Elements<PageSetup>().LastOrDefault());
            contents.Add(XLWSContents.HeaderFooter, opWorksheet.Elements<HeaderFooter>().LastOrDefault());
            contents.Add(XLWSContents.RowBreaks, opWorksheet.Elements<RowBreaks>().LastOrDefault());
            contents.Add(XLWSContents.ColumnBreaks, opWorksheet.Elements<ColumnBreaks>().LastOrDefault());
            contents.Add(XLWSContents.CustomProperties, opWorksheet.Elements<CustomProperties>().LastOrDefault());
            contents.Add(XLWSContents.CellWatches, opWorksheet.Elements<CellWatches>().LastOrDefault());
            contents.Add(XLWSContents.IgnoredErrors, opWorksheet.Elements<IgnoredErrors>().LastOrDefault());
            //contents.Add(XLWSContents.SmartTags, opWorksheet.Elements<SmartTags>().LastOrDefault());
            contents.Add(XLWSContents.Drawing, opWorksheet.Elements<Drawing>().LastOrDefault());
            contents.Add(XLWSContents.LegacyDrawing, opWorksheet.Elements<LegacyDrawing>().LastOrDefault());
            contents.Add(XLWSContents.LegacyDrawingHeaderFooter, opWorksheet.Elements<LegacyDrawingHeaderFooter>().LastOrDefault());
            contents.Add(XLWSContents.DrawingHeaderFooter, opWorksheet.Elements<DrawingHeaderFooter>().LastOrDefault());
            contents.Add(XLWSContents.Picture, opWorksheet.Elements<Picture>().LastOrDefault());
            contents.Add(XLWSContents.OleObjects, opWorksheet.Elements<OleObjects>().LastOrDefault());
            contents.Add(XLWSContents.Controls, opWorksheet.Elements<Controls>().LastOrDefault());
            contents.Add(XLWSContents.AlternateContent, opWorksheet.Elements<AlternateContent>().LastOrDefault());
            contents.Add(XLWSContents.WebPublishItems, opWorksheet.Elements<WebPublishItems>().LastOrDefault());
            contents.Add(XLWSContents.TableParts, opWorksheet.Elements<TableParts>().LastOrDefault());
            contents.Add(XLWSContents.WorksheetExtensionList, opWorksheet.Elements<WorksheetExtensionList>().LastOrDefault());
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
