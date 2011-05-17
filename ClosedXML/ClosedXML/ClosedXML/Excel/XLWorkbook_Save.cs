using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Op = DocumentFormat.OpenXml.CustomProperties;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Globalization;
using DRW = System.Drawing;

namespace ClosedXML.Excel
{
    public partial class XLWorkbook
    {
        private List<KeyValuePair<XLFillPatternValues, PatternValues>> fillPatternValues = new List<KeyValuePair<XLFillPatternValues, PatternValues>>();
        private List<KeyValuePair<XLAlignmentHorizontalValues, HorizontalAlignmentValues>> alignmentHorizontalValues = new List<KeyValuePair<XLAlignmentHorizontalValues, HorizontalAlignmentValues>>();
        private List<KeyValuePair<XLAlignmentVerticalValues, VerticalAlignmentValues>> alignmentVerticalValues = new List<KeyValuePair<XLAlignmentVerticalValues, VerticalAlignmentValues>>();
        private List<KeyValuePair<XLBorderStyleValues, BorderStyleValues>> borderStyleValues = new List<KeyValuePair<XLBorderStyleValues, BorderStyleValues>>();
        private List<KeyValuePair<XLFontUnderlineValues, UnderlineValues>> underlineValuesList = new List<KeyValuePair<XLFontUnderlineValues, UnderlineValues>>();
        private List<KeyValuePair<XLFontVerticalTextAlignmentValues, VerticalAlignmentRunValues>> fontVerticalTextAlignmentValues = new List<KeyValuePair<XLFontVerticalTextAlignmentValues, VerticalAlignmentRunValues>>();
        private List<KeyValuePair<XLPageOrderValues, PageOrderValues>> pageOrderValues = new List<KeyValuePair<XLPageOrderValues, PageOrderValues>>();
        private List<KeyValuePair<XLPageOrientation, OrientationValues>> pageOrientationValues = new List<KeyValuePair<XLPageOrientation, OrientationValues>>();
        private List<KeyValuePair<XLShowCommentsValues, CellCommentsValues>> showCommentsValues = new List<KeyValuePair<XLShowCommentsValues, CellCommentsValues>>();
        private List<KeyValuePair<XLPrintErrorValues, PrintErrorValues>> printErrorValues = new List<KeyValuePair<XLPrintErrorValues, PrintErrorValues>>();
        private List<KeyValuePair<XLCalculateMode, CalculateModeValues>> calculateModeValues = new List<KeyValuePair<XLCalculateMode, CalculateModeValues>>();
        private List<KeyValuePair<XLReferenceStyle, ReferenceModeValues>> referenceModeValues = new List<KeyValuePair<XLReferenceStyle, ReferenceModeValues>>();
        private List<KeyValuePair<XLAlignmentReadingOrderValues, UInt32>> alignmentReadingOrderValues = new List<KeyValuePair<XLAlignmentReadingOrderValues, UInt32>>();
        private List<KeyValuePair<XLTotalsRowFunction, TotalsRowFunctionValues>> totalsRowFunctionValues = new List<KeyValuePair<XLTotalsRowFunction, TotalsRowFunctionValues>>();
        private List<KeyValuePair<XLAllowedValues, DataValidationValues>> dataValidationValues = new List<KeyValuePair<XLAllowedValues, DataValidationValues>>();
        private List<KeyValuePair<XLErrorStyle, DataValidationErrorStyleValues>> dataValidationErrorStyleValues = new List<KeyValuePair<XLErrorStyle, DataValidationErrorStyleValues>>();
        private List<KeyValuePair<XLOperator, DataValidationOperatorValues>> dataValidationOperatorValues = new List<KeyValuePair<XLOperator, DataValidationOperatorValues>>();
        private List<KeyValuePair<XLWorksheetVisibility, SheetStateValues>> sheetStateValues = new List<KeyValuePair<XLWorksheetVisibility, SheetStateValues>>();

        private Boolean populated = false;
        private void PopulateEnums()
        {
            if (!populated)
            {
                PopulateFillPatternValues();
                PopulateAlignmentHorizontalValues();
                PopulateAlignmentVerticalValues();
                PupulateBorderStyleValues();
                PopulateUnderlineValues();
                PopulateFontVerticalTextAlignmentValues();
                PopulatePageOrderValues();
                PopulatePageOrientationValues();
                PopulateShowCommentsValues();
                PopulatePrintErrorValues();
                PopulateCalculateModeValues();
                PopulateReferenceModeValues();
                PopulateAlignmentReadingOrderValues();
                PopulateTotalsRowFunctionValues();
                PopulateDataValidationValues();
                PopulateDataValidationErrorStyleValues();
                PopulateDataValidationOperatorValues();
                PopulateSheetStateValues();
                populated = true;
            }
        }

        private enum RelType { General, Workbook, Worksheet, Drawing }
        private class RelIdGenerator
        {
            private Dictionary<RelType, List<String>> relIds = new Dictionary<RelType, List<String>>();
            public String GetNext(RelType relType)
            {
                if (!relIds.ContainsKey(relType))
                    relIds.Add(relType, new List<String>());

                Int32 id = 1;
                while (true)
                {
                    String relId = String.Format("rId{0}", id);
                    if (!relIds[relType].Contains(relId))
                    {
                        relIds[relType].Add(relId);
                        return relId;
                    }
                    id++;
                }
            }
            public void AddValues(List<String> values, RelType relType)
            {
                if (!relIds.ContainsKey(relType))
                    relIds.Add(relType, new List<String>());
                relIds[relType].AddRange(values);
            }
        }

        private Dictionary<String, UInt32> sharedStrings;
        private Dictionary<IXLStyle, StyleInfo> sharedStyles;
        private struct FontInfo { public UInt32 FontId; public IXLFont Font; };
        private struct FillInfo { public UInt32 FillId; public IXLFill Fill; };
        private struct BorderInfo { public UInt32 BorderId; public IXLBorder Border; };
        private struct NumberFormatInfo { public Int32 NumberFormatId; public IXLNumberFormat NumberFormat; };


        private struct StyleInfo
        {
            public UInt32 StyleId;
            public UInt32 FontId;
            public UInt32 FillId;
            public UInt32 BorderId;
            public Int32 NumberFormatId;
            public IXLStyle Style;
        };

        private CellValues GetCellValue(XLCell xlCell)
        {
            var xlCellValue = xlCell.DataType;
            switch (xlCellValue)
            {
                case XLCellValues.Boolean: return CellValues.Boolean;
                case XLCellValues.DateTime: return CellValues.Date;
                case XLCellValues.Number: return CellValues.Number;
                case XLCellValues.Text:
                    {
                        if (xlCell.ShareString)
                            return CellValues.SharedString;
                        else
                            return CellValues.InlineString;
                    }
                case XLCellValues.TimeSpan: return CellValues.Number;
                default: throw new NotImplementedException();
            }
        }

        #region Pouplate Values
        private void PopulateUnderlineValues()
        {

            underlineValuesList.Add(new KeyValuePair<XLFontUnderlineValues, UnderlineValues>(XLFontUnderlineValues.Double, UnderlineValues.Double));
            underlineValuesList.Add(new KeyValuePair<XLFontUnderlineValues, UnderlineValues>(XLFontUnderlineValues.DoubleAccounting, UnderlineValues.DoubleAccounting));
            underlineValuesList.Add(new KeyValuePair<XLFontUnderlineValues, UnderlineValues>(XLFontUnderlineValues.None, UnderlineValues.None));
            underlineValuesList.Add(new KeyValuePair<XLFontUnderlineValues, UnderlineValues>(XLFontUnderlineValues.Single, UnderlineValues.Single));
            underlineValuesList.Add(new KeyValuePair<XLFontUnderlineValues, UnderlineValues>(XLFontUnderlineValues.SingleAccounting, UnderlineValues.SingleAccounting));
        }

        private void PopulatePageOrientationValues()
        {
            pageOrientationValues.Add(new KeyValuePair<XLPageOrientation, OrientationValues>(XLPageOrientation.Default, OrientationValues.Default));
            pageOrientationValues.Add(new KeyValuePair<XLPageOrientation, OrientationValues>(XLPageOrientation.Landscape, OrientationValues.Landscape));
            pageOrientationValues.Add(new KeyValuePair<XLPageOrientation, OrientationValues>(XLPageOrientation.Portrait, OrientationValues.Portrait));
        }

        private void PopulateFontVerticalTextAlignmentValues()
        {
            fontVerticalTextAlignmentValues.Add(new KeyValuePair<XLFontVerticalTextAlignmentValues, VerticalAlignmentRunValues>(XLFontVerticalTextAlignmentValues.Baseline, VerticalAlignmentRunValues.Baseline));
            fontVerticalTextAlignmentValues.Add(new KeyValuePair<XLFontVerticalTextAlignmentValues, VerticalAlignmentRunValues>(XLFontVerticalTextAlignmentValues.Subscript, VerticalAlignmentRunValues.Subscript));
            fontVerticalTextAlignmentValues.Add(new KeyValuePair<XLFontVerticalTextAlignmentValues, VerticalAlignmentRunValues>(XLFontVerticalTextAlignmentValues.Superscript, VerticalAlignmentRunValues.Superscript));
        }

        private void PopulateFillPatternValues()
        { 
                fillPatternValues.Add(new KeyValuePair<XLFillPatternValues,PatternValues>(XLFillPatternValues.DarkDown, PatternValues.DarkDown));
                fillPatternValues.Add(new KeyValuePair<XLFillPatternValues,PatternValues>(XLFillPatternValues.DarkGray, PatternValues.DarkGray));
                fillPatternValues.Add(new KeyValuePair<XLFillPatternValues,PatternValues>(XLFillPatternValues.DarkGrid, PatternValues.DarkGrid));
                fillPatternValues.Add(new KeyValuePair<XLFillPatternValues,PatternValues>(XLFillPatternValues.DarkHorizontal, PatternValues.DarkHorizontal));
                fillPatternValues.Add(new KeyValuePair<XLFillPatternValues,PatternValues>(XLFillPatternValues.DarkTrellis, PatternValues.DarkTrellis));
                fillPatternValues.Add(new KeyValuePair<XLFillPatternValues,PatternValues>(XLFillPatternValues.DarkUp, PatternValues.DarkUp));
                fillPatternValues.Add(new KeyValuePair<XLFillPatternValues,PatternValues>(XLFillPatternValues.DarkVertical, PatternValues.DarkVertical));
                fillPatternValues.Add(new KeyValuePair<XLFillPatternValues,PatternValues>(XLFillPatternValues.Gray0625, PatternValues.Gray0625));
                fillPatternValues.Add(new KeyValuePair<XLFillPatternValues,PatternValues>(XLFillPatternValues.Gray125, PatternValues.Gray125));
                fillPatternValues.Add(new KeyValuePair<XLFillPatternValues,PatternValues>(XLFillPatternValues.LightDown, PatternValues.LightDown));
                fillPatternValues.Add(new KeyValuePair<XLFillPatternValues,PatternValues>(XLFillPatternValues.LightGray, PatternValues.LightGray));
                fillPatternValues.Add(new KeyValuePair<XLFillPatternValues,PatternValues>(XLFillPatternValues.LightGrid, PatternValues.LightGrid));
                fillPatternValues.Add(new KeyValuePair<XLFillPatternValues,PatternValues>(XLFillPatternValues.LightHorizontal, PatternValues.LightHorizontal));
                fillPatternValues.Add(new KeyValuePair<XLFillPatternValues,PatternValues>(XLFillPatternValues.LightTrellis, PatternValues.LightTrellis));
                fillPatternValues.Add(new KeyValuePair<XLFillPatternValues,PatternValues>(XLFillPatternValues.LightUp, PatternValues.LightUp));
                fillPatternValues.Add(new KeyValuePair<XLFillPatternValues,PatternValues>(XLFillPatternValues.LightVertical, PatternValues.LightVertical));
                fillPatternValues.Add(new KeyValuePair<XLFillPatternValues,PatternValues>(XLFillPatternValues.MediumGray, PatternValues.MediumGray));
                fillPatternValues.Add(new KeyValuePair<XLFillPatternValues,PatternValues>(XLFillPatternValues.None, PatternValues.None));
                fillPatternValues.Add(new KeyValuePair<XLFillPatternValues,PatternValues>(XLFillPatternValues.Solid, PatternValues.Solid));
        }

        private void PupulateBorderStyleValues()
        {

            borderStyleValues.Add(new KeyValuePair<XLBorderStyleValues, BorderStyleValues>(XLBorderStyleValues.DashDot, BorderStyleValues.DashDot));
            borderStyleValues.Add(new KeyValuePair<XLBorderStyleValues, BorderStyleValues>(XLBorderStyleValues.DashDotDot, BorderStyleValues.DashDotDot));
            borderStyleValues.Add(new KeyValuePair<XLBorderStyleValues, BorderStyleValues>(XLBorderStyleValues.Dashed, BorderStyleValues.Dashed));
            borderStyleValues.Add(new KeyValuePair<XLBorderStyleValues, BorderStyleValues>(XLBorderStyleValues.Dotted, BorderStyleValues.Dotted));
            borderStyleValues.Add(new KeyValuePair<XLBorderStyleValues, BorderStyleValues>(XLBorderStyleValues.Double, BorderStyleValues.Double));
            borderStyleValues.Add(new KeyValuePair<XLBorderStyleValues, BorderStyleValues>(XLBorderStyleValues.Hair, BorderStyleValues.Hair));
            borderStyleValues.Add(new KeyValuePair<XLBorderStyleValues, BorderStyleValues>(XLBorderStyleValues.Medium, BorderStyleValues.Medium));
            borderStyleValues.Add(new KeyValuePair<XLBorderStyleValues, BorderStyleValues>(XLBorderStyleValues.MediumDashDot, BorderStyleValues.MediumDashDot));
            borderStyleValues.Add(new KeyValuePair<XLBorderStyleValues, BorderStyleValues>(XLBorderStyleValues.MediumDashDotDot, BorderStyleValues.MediumDashDotDot));
            borderStyleValues.Add(new KeyValuePair<XLBorderStyleValues, BorderStyleValues>(XLBorderStyleValues.MediumDashed, BorderStyleValues.MediumDashed));
            borderStyleValues.Add(new KeyValuePair<XLBorderStyleValues, BorderStyleValues>(XLBorderStyleValues.None, BorderStyleValues.None));
            borderStyleValues.Add(new KeyValuePair<XLBorderStyleValues, BorderStyleValues>(XLBorderStyleValues.SlantDashDot, BorderStyleValues.SlantDashDot));
            borderStyleValues.Add(new KeyValuePair<XLBorderStyleValues, BorderStyleValues>(XLBorderStyleValues.Thick, BorderStyleValues.Thick));
            borderStyleValues.Add(new KeyValuePair<XLBorderStyleValues, BorderStyleValues>(XLBorderStyleValues.Thin, BorderStyleValues.Thin));

        }

        private void PopulateAlignmentHorizontalValues()
        {
                alignmentHorizontalValues.Add(new KeyValuePair<XLAlignmentHorizontalValues,HorizontalAlignmentValues>(XLAlignmentHorizontalValues.Center, HorizontalAlignmentValues.Center));
                alignmentHorizontalValues.Add(new KeyValuePair<XLAlignmentHorizontalValues,HorizontalAlignmentValues>(XLAlignmentHorizontalValues.CenterContinuous, HorizontalAlignmentValues.CenterContinuous));
                alignmentHorizontalValues.Add(new KeyValuePair<XLAlignmentHorizontalValues,HorizontalAlignmentValues>(XLAlignmentHorizontalValues.Distributed, HorizontalAlignmentValues.Distributed));
                alignmentHorizontalValues.Add(new KeyValuePair<XLAlignmentHorizontalValues,HorizontalAlignmentValues>(XLAlignmentHorizontalValues.Fill, HorizontalAlignmentValues.Fill));
                alignmentHorizontalValues.Add(new KeyValuePair<XLAlignmentHorizontalValues,HorizontalAlignmentValues>(XLAlignmentHorizontalValues.General, HorizontalAlignmentValues.General));
                alignmentHorizontalValues.Add(new KeyValuePair<XLAlignmentHorizontalValues,HorizontalAlignmentValues>(XLAlignmentHorizontalValues.Justify, HorizontalAlignmentValues.Justify));
                alignmentHorizontalValues.Add(new KeyValuePair<XLAlignmentHorizontalValues,HorizontalAlignmentValues>(XLAlignmentHorizontalValues.Left, HorizontalAlignmentValues.Left));
                alignmentHorizontalValues.Add(new KeyValuePair<XLAlignmentHorizontalValues,HorizontalAlignmentValues>(XLAlignmentHorizontalValues.Right, HorizontalAlignmentValues.Right));
        }

        private void PopulateAlignmentVerticalValues()
        {

            alignmentVerticalValues.Add(new KeyValuePair<XLAlignmentVerticalValues, VerticalAlignmentValues>(XLAlignmentVerticalValues.Bottom, VerticalAlignmentValues.Bottom));
            alignmentVerticalValues.Add(new KeyValuePair<XLAlignmentVerticalValues, VerticalAlignmentValues>(XLAlignmentVerticalValues.Center, VerticalAlignmentValues.Center));
            alignmentVerticalValues.Add(new KeyValuePair<XLAlignmentVerticalValues, VerticalAlignmentValues>(XLAlignmentVerticalValues.Distributed, VerticalAlignmentValues.Distributed));
            alignmentVerticalValues.Add(new KeyValuePair<XLAlignmentVerticalValues, VerticalAlignmentValues>(XLAlignmentVerticalValues.Justify, VerticalAlignmentValues.Justify));
            alignmentVerticalValues.Add(new KeyValuePair<XLAlignmentVerticalValues, VerticalAlignmentValues>(XLAlignmentVerticalValues.Top, VerticalAlignmentValues.Top));
        }

        private void PopulatePageOrderValues()
        {
            pageOrderValues.Add(new KeyValuePair<XLPageOrderValues, PageOrderValues>(XLPageOrderValues.DownThenOver, PageOrderValues.DownThenOver));
            pageOrderValues.Add(new KeyValuePair<XLPageOrderValues, PageOrderValues>(XLPageOrderValues.OverThenDown, PageOrderValues.OverThenDown));
        }

        private void PopulateShowCommentsValues()
        {
            showCommentsValues.Add(new KeyValuePair<XLShowCommentsValues, CellCommentsValues>(XLShowCommentsValues.AsDisplayed, CellCommentsValues.AsDisplayed));
            showCommentsValues.Add(new KeyValuePair<XLShowCommentsValues, CellCommentsValues>(XLShowCommentsValues.AtEnd, CellCommentsValues.AtEnd));
            showCommentsValues.Add(new KeyValuePair<XLShowCommentsValues, CellCommentsValues>(XLShowCommentsValues.None, CellCommentsValues.None));
        }

        private void PopulatePrintErrorValues()
        {
            printErrorValues.Add(new KeyValuePair<XLPrintErrorValues, PrintErrorValues>(XLPrintErrorValues.Blank, PrintErrorValues.Blank));
            printErrorValues.Add(new KeyValuePair<XLPrintErrorValues, PrintErrorValues>(XLPrintErrorValues.Dash, PrintErrorValues.Dash));
            printErrorValues.Add(new KeyValuePair<XLPrintErrorValues, PrintErrorValues>(XLPrintErrorValues.Displayed, PrintErrorValues.Displayed));
            printErrorValues.Add(new KeyValuePair<XLPrintErrorValues, PrintErrorValues>(XLPrintErrorValues.NA, PrintErrorValues.NA));
        }

        private void PopulateCalculateModeValues()
        {
            calculateModeValues.Add(new KeyValuePair<XLCalculateMode, CalculateModeValues>(XLCalculateMode.Auto, CalculateModeValues.Auto)) ;
            calculateModeValues.Add(new KeyValuePair<XLCalculateMode, CalculateModeValues>(XLCalculateMode.AutoNoTable, CalculateModeValues.AutoNoTable));
            calculateModeValues.Add(new KeyValuePair<XLCalculateMode, CalculateModeValues>(XLCalculateMode.Manual, CalculateModeValues.Manual));
        }

        private void PopulateReferenceModeValues()
        {
            referenceModeValues.Add(new KeyValuePair<XLReferenceStyle, ReferenceModeValues>(XLReferenceStyle.R1C1, ReferenceModeValues.R1C1));
            referenceModeValues.Add(new KeyValuePair<XLReferenceStyle, ReferenceModeValues>(XLReferenceStyle.A1, ReferenceModeValues.A1));
        }

        private void PopulateAlignmentReadingOrderValues()
        {
            alignmentReadingOrderValues.Add(new KeyValuePair<XLAlignmentReadingOrderValues, uint>(XLAlignmentReadingOrderValues.ContextDependent, 0));
            alignmentReadingOrderValues.Add(new KeyValuePair<XLAlignmentReadingOrderValues, uint>(XLAlignmentReadingOrderValues.LeftToRight, 1));
            alignmentReadingOrderValues.Add(new KeyValuePair<XLAlignmentReadingOrderValues, uint>(XLAlignmentReadingOrderValues.RightToLeft, 2));
        }

        private void PopulateTotalsRowFunctionValues()
        {
            totalsRowFunctionValues.Add(new KeyValuePair<XLTotalsRowFunction, TotalsRowFunctionValues>(XLTotalsRowFunction.None, TotalsRowFunctionValues.None));
            totalsRowFunctionValues.Add(new KeyValuePair<XLTotalsRowFunction, TotalsRowFunctionValues>(XLTotalsRowFunction.Sum, TotalsRowFunctionValues.Sum));
            totalsRowFunctionValues.Add(new KeyValuePair<XLTotalsRowFunction, TotalsRowFunctionValues>(XLTotalsRowFunction.Minimum, TotalsRowFunctionValues.Minimum));
            totalsRowFunctionValues.Add(new KeyValuePair<XLTotalsRowFunction, TotalsRowFunctionValues>(XLTotalsRowFunction.Maximum, TotalsRowFunctionValues.Maximum));
            totalsRowFunctionValues.Add(new KeyValuePair<XLTotalsRowFunction, TotalsRowFunctionValues>(XLTotalsRowFunction.Average, TotalsRowFunctionValues.Average));
            totalsRowFunctionValues.Add(new KeyValuePair<XLTotalsRowFunction, TotalsRowFunctionValues>(XLTotalsRowFunction.Count, TotalsRowFunctionValues.Count));
            totalsRowFunctionValues.Add(new KeyValuePair<XLTotalsRowFunction, TotalsRowFunctionValues>(XLTotalsRowFunction.CountNumbers, TotalsRowFunctionValues.CountNumbers));
            totalsRowFunctionValues.Add(new KeyValuePair<XLTotalsRowFunction, TotalsRowFunctionValues>(XLTotalsRowFunction.StandardDeviation, TotalsRowFunctionValues.StandardDeviation));
            totalsRowFunctionValues.Add(new KeyValuePair<XLTotalsRowFunction, TotalsRowFunctionValues>(XLTotalsRowFunction.Variance, TotalsRowFunctionValues.Variance));
            totalsRowFunctionValues.Add(new KeyValuePair<XLTotalsRowFunction, TotalsRowFunctionValues>(XLTotalsRowFunction.Custom, TotalsRowFunctionValues.Custom));
        }

        private void PopulateDataValidationValues()
        {
            dataValidationValues.Add(new KeyValuePair<XLAllowedValues, DataValidationValues>(XLAllowedValues.AnyValue, DataValidationValues.None));
            dataValidationValues.Add(new KeyValuePair<XLAllowedValues, DataValidationValues>(XLAllowedValues.Custom, DataValidationValues.Custom));
            dataValidationValues.Add(new KeyValuePair<XLAllowedValues, DataValidationValues>(XLAllowedValues.Date, DataValidationValues.Date));
            dataValidationValues.Add(new KeyValuePair<XLAllowedValues, DataValidationValues>(XLAllowedValues.Decimal, DataValidationValues.Decimal));
            dataValidationValues.Add(new KeyValuePair<XLAllowedValues, DataValidationValues>(XLAllowedValues.List, DataValidationValues.List));
            dataValidationValues.Add(new KeyValuePair<XLAllowedValues, DataValidationValues>(XLAllowedValues.TextLength, DataValidationValues.TextLength));
            dataValidationValues.Add(new KeyValuePair<XLAllowedValues, DataValidationValues>(XLAllowedValues.Time, DataValidationValues.Time));
            dataValidationValues.Add(new KeyValuePair<XLAllowedValues, DataValidationValues>(XLAllowedValues.WholeNumber, DataValidationValues.Whole));
        }

        private void PopulateDataValidationErrorStyleValues()
        {
            dataValidationErrorStyleValues.Add(new KeyValuePair<XLErrorStyle, DataValidationErrorStyleValues>(XLErrorStyle.Information, DataValidationErrorStyleValues.Information));
            dataValidationErrorStyleValues.Add(new KeyValuePair<XLErrorStyle, DataValidationErrorStyleValues>(XLErrorStyle.Warning, DataValidationErrorStyleValues.Warning));
            dataValidationErrorStyleValues.Add(new KeyValuePair<XLErrorStyle, DataValidationErrorStyleValues>(XLErrorStyle.Stop, DataValidationErrorStyleValues.Stop));
        }

        private void PopulateDataValidationOperatorValues()
        {
            dataValidationOperatorValues.Add(new KeyValuePair<XLOperator, DataValidationOperatorValues>(XLOperator.Between, DataValidationOperatorValues.Between));
            dataValidationOperatorValues.Add(new KeyValuePair<XLOperator, DataValidationOperatorValues>(XLOperator.EqualOrGreaterThan , DataValidationOperatorValues.GreaterThanOrEqual ));
            dataValidationOperatorValues.Add(new KeyValuePair<XLOperator, DataValidationOperatorValues>(XLOperator.EqualOrLessThan, DataValidationOperatorValues.LessThanOrEqual));
            dataValidationOperatorValues.Add(new KeyValuePair<XLOperator, DataValidationOperatorValues>(XLOperator.EqualTo, DataValidationOperatorValues.Equal));
            dataValidationOperatorValues.Add(new KeyValuePair<XLOperator, DataValidationOperatorValues>(XLOperator.GreaterThan, DataValidationOperatorValues.GreaterThan));
            dataValidationOperatorValues.Add(new KeyValuePair<XLOperator, DataValidationOperatorValues>(XLOperator.LessThan, DataValidationOperatorValues.LessThan));
            dataValidationOperatorValues.Add(new KeyValuePair<XLOperator, DataValidationOperatorValues>(XLOperator.NotBetween, DataValidationOperatorValues.NotBetween));
            dataValidationOperatorValues.Add(new KeyValuePair<XLOperator, DataValidationOperatorValues>(XLOperator.NotEqualTo, DataValidationOperatorValues.NotEqual));
        }

        private void PopulateSheetStateValues()
        {
            sheetStateValues.Add(new KeyValuePair<XLWorksheetVisibility, SheetStateValues>(XLWorksheetVisibility.Visible, SheetStateValues.Visible));
            sheetStateValues.Add(new KeyValuePair<XLWorksheetVisibility, SheetStateValues>(XLWorksheetVisibility.Hidden, SheetStateValues.Hidden));
            sheetStateValues.Add(new KeyValuePair<XLWorksheetVisibility, SheetStateValues>(XLWorksheetVisibility.VeryHidden, SheetStateValues.VeryHidden));
        }

        #endregion

        private void CreatePackage(String filePath)
        {
            SpreadsheetDocument package;
            if (File.Exists(filePath))
                package = SpreadsheetDocument.Open(filePath, true);
            else
                package = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);

            using (package)
            {
                CreateParts(package);
            }
        }

        private void CreatePackage(Stream stream, Boolean newStream)
        {
            SpreadsheetDocument package;
            if (newStream)
                package = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
            else
                package = SpreadsheetDocument.Open(stream, true);

            using (package)
            {
                CreateParts(package);
            }
        }

        // Adds child parts and generates content of the specified part.
        private RelIdGenerator relId;
        private void CreateParts(SpreadsheetDocument document)
        {
            relId = new RelIdGenerator();
            sharedStrings = new Dictionary<String, UInt32>();
            sharedStyles = new Dictionary<IXLStyle, StyleInfo>();
            tableNames = new HashSet<String>();
            tableId = 0;

            WorkbookPart workbookPart;
            if (document.WorkbookPart == null)
                workbookPart = document.AddWorkbookPart();
            else
                workbookPart = document.WorkbookPart;
            
            var worksheets = (XLWorksheets)Worksheets;
            var partsToRemove = workbookPart.Parts.Where(s => worksheets.Deleted.Contains(s.RelationshipId)).ToList();
            partsToRemove.ForEach(s => workbookPart.DeletePart(s.OpenXmlPart));
            relId.AddValues(workbookPart.Parts.Select(p=>p.RelationshipId).ToList(), RelType.Workbook);

            var modifiedSheetNames = Worksheets.Select(w => w.Name.ToLower()).ToList();

            List<String> existingSheetNames;
            if (workbookPart.Workbook != null && workbookPart.Workbook.Sheets != null)
                existingSheetNames = workbookPart.Workbook.Sheets.Elements<Sheet>().Select(s => s.Name.Value.ToLower()).ToList();
            else
                existingSheetNames = new List<String>();

            var allSheetNames = existingSheetNames.Union(modifiedSheetNames);

            ExtendedFilePropertiesPart extendedFilePropertiesPart;
            if (document.ExtendedFilePropertiesPart == null)
                extendedFilePropertiesPart = document.AddNewPart<ExtendedFilePropertiesPart>(relId.GetNext(RelType.Workbook));
            else
                extendedFilePropertiesPart = document.ExtendedFilePropertiesPart;

            GenerateExtendedFilePropertiesPartContent(extendedFilePropertiesPart, workbookPart);

            GenerateWorkbookPartContent(workbookPart);
   
            SharedStringTablePart sharedStringTablePart;
            if (workbookPart.SharedStringTablePart == null)
                sharedStringTablePart = workbookPart.AddNewPart<SharedStringTablePart>(relId.GetNext(RelType.Workbook));
            else
                sharedStringTablePart = workbookPart.SharedStringTablePart;
             
            GenerateSharedStringTablePartContent(sharedStringTablePart);

            WorkbookStylesPart workbookStylesPart;
            if (workbookPart.WorkbookStylesPart == null)
                workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>(relId.GetNext(RelType.Workbook));
            else
                workbookStylesPart = workbookPart.WorkbookStylesPart;

            GenerateWorkbookStylesPartContent(workbookStylesPart);

            foreach (var worksheet in Worksheets.Cast<XLWorksheet>().OrderBy(w=>w.Position))
            {
                WorksheetPart worksheetPart;
                if (workbookPart.Parts.Where(p => p.RelationshipId == worksheet.RelId).Any())
                {
                    worksheetPart = (WorksheetPart)workbookPart.GetPartById(worksheet.RelId);
                    var wsPartsToRemove = worksheetPart.TableDefinitionParts.ToList();
                    wsPartsToRemove.ForEach(tdp=>worksheetPart.DeletePart(tdp));
                }
                else
                {
                    worksheetPart = workbookPart.AddNewPart<WorksheetPart>(worksheet.RelId);
                }

                GenerateWorksheetPartContent(worksheetPart, worksheet);

                //DrawingsPart drawingsPart = worksheetPart.AddNewPart<DrawingsPart>("rId1");
                //GenerateDrawingsPartContent(drawingsPart, worksheet);

                //foreach (var chart in worksheet.Charts)
                //{
                //    ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>("rId1");
                //    GenerateChartPartContent(chartPart, (XLChart)chart);
                //}
            }


            GenerateCalculationChainPartContent(workbookPart);

            if (workbookPart.ThemePart == null)
            {
                ThemePart themePart = workbookPart.AddNewPart<ThemePart>(relId.GetNext(RelType.Workbook));
                GenerateThemePartContent(themePart);
            }

            if (CustomProperties.Count() > 0)
            {
                document.GetPartsOfType<CustomFilePropertiesPart>().ToList().ForEach(p => document.DeletePart(p));
                CustomFilePropertiesPart customFilePropertiesPart = document.AddNewPart<CustomFilePropertiesPart>(relId.GetNext(RelType.Workbook));
                
                GenerateCustomFilePropertiesPartContent(customFilePropertiesPart);
            }
            SetPackageProperties(document);
        }

        private void GenerateTables(XLWorksheet worksheet, WorksheetPart worksheetPart)
        {
            worksheetPart.Worksheet.RemoveAllChildren<TablePart>();
            if (worksheet.Tables.Count() > 0)
            {
                foreach (var table in worksheet.Tables)
                {
                    String tableRelId = relId.GetNext(RelType.Workbook);
                    var xlTable = (XLTable)table;
                    xlTable.RelId = tableRelId;
                    TableDefinitionPart tableDefinitionPart = worksheetPart.AddNewPart<TableDefinitionPart>(tableRelId);
                    GenerateTableDefinitionPartContent(tableDefinitionPart, xlTable);
                }
            }
        }

        private void GenerateExtendedFilePropertiesPartContent(ExtendedFilePropertiesPart extendedFilePropertiesPart, WorkbookPart workbookPart)
        {
            //if (extendedFilePropertiesPart.Properties.NamespaceDeclarations.Contains(new KeyValuePair<string,string>(
            Ap.Properties properties;
            if (extendedFilePropertiesPart.Properties == null)
                extendedFilePropertiesPart.Properties = new Ap.Properties();

            properties = extendedFilePropertiesPart.Properties;
            if (!properties.NamespaceDeclarations.Contains(new KeyValuePair<string,string>("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")))
                properties.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");

            if (properties.Application == null)
                properties.Append(new Ap.Application() { Text = "Microsoft Excel" });

            if (properties.DocumentSecurity == null)
                properties.Append(new Ap.DocumentSecurity() { Text = "0" });

            if (properties.ScaleCrop == null)
                properties.Append(new Ap.ScaleCrop() { Text = "false" });

            if (properties.HeadingPairs == null)
                properties.HeadingPairs = new Ap.HeadingPairs();

            if (properties.TitlesOfParts == null)
                properties.TitlesOfParts = new Ap.TitlesOfParts();

            properties.HeadingPairs.VTVector = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Variant};

            properties.TitlesOfParts.VTVector = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Lpstr };

            Vt.VTVector vTVector_One;
            vTVector_One = properties.HeadingPairs.VTVector;

            Vt.VTVector vTVector_Two;
            vTVector_Two = properties.TitlesOfParts.VTVector;


            var modifiedWorksheets = Worksheets.Select(w => new { w.Name, Order = w.Position });
            var modifiedNamedRanges = GetModifiedNamedRanges();

            InsertOnVTVector(vTVector_One, "Worksheets", 0, modifiedWorksheets.Count().ToString());
            InsertOnVTVector(vTVector_One, "Named Ranges", 2, (modifiedNamedRanges.Count()).ToString());

            vTVector_Two.Size = (UInt32)(modifiedNamedRanges.Count() + modifiedWorksheets.Count());

            foreach (var w in modifiedWorksheets.OrderBy(w=>w.Order))
            {
                Vt.VTLPSTR vTLPSTR3 = new Vt.VTLPSTR() { Text = w.Name };
                vTVector_Two.Append(vTLPSTR3);
            }

            foreach (var nr in modifiedNamedRanges)
            {
                Vt.VTLPSTR vTLPSTR7 = new Vt.VTLPSTR() { Text = nr };
                vTVector_Two.Append(vTLPSTR7);
            }

            if (Properties.Manager != null)
            {
                if (!StringExtensions.IsNullOrWhiteSpace(Properties.Manager))
                {
                    if (properties.Manager == null)
                        properties.Manager = new Ap.Manager();

                    properties.Manager.Text = Properties.Manager;
                }
                else
                {
                    properties.Manager = null;
                }
            }

            if (Properties.Company != null)
            {
                if (!StringExtensions.IsNullOrWhiteSpace(Properties.Company))
                {
                    if (properties.Company == null)
                        properties.Company = new Ap.Company();

                    properties.Company.Text = Properties.Company;
                }
                else
                {
                    properties = null;
                }
            }
        }

        private void InsertOnVTVector(Vt.VTVector vTVector, String property, Int32 index, String text)
        {
            var m = from e1 in vTVector.Elements<Vt.Variant>()
                    where e1.Elements<Vt.VTLPSTR>().Any(e2 => e2.Text == property)
                    select e1;
            if (m.Count() == 0)
            {
                if (vTVector.Size == null)
                    vTVector.Size = new UInt32Value(0U);

                vTVector.Size += 2U;
                Vt.Variant variant1 = new Vt.Variant();
                Vt.VTLPSTR vTLPSTR1 = new Vt.VTLPSTR() { Text = property };
                variant1.Append(vTLPSTR1);
                vTVector.InsertAt<Vt.Variant>(variant1, index);

                Vt.Variant variant2 = new Vt.Variant();
                Vt.VTInt32 vTInt321 = new Vt.VTInt32();
                variant2.Append(vTInt321);
                vTVector.InsertAt<Vt.Variant>(variant2, index + 1);
            }

            Int32 targetIndex = 0;
            foreach (var e in vTVector.Elements<Vt.Variant>())
            {
                if (e.Elements<Vt.VTLPSTR>().Any(e2 => e2.Text == property))
                {
                    vTVector.ElementAt(targetIndex + 1).GetFirstChild<Vt.VTInt32>().Text = text;
                    break;
                }
                targetIndex++;
            }
        }

        private List<String> GetExistingWorksheets(WorkbookPart workbookPart)
        {
            if (workbookPart != null && workbookPart.Workbook != null && workbookPart.Workbook.Sheets != null)
                return workbookPart.Workbook.Sheets.Select(s=>((Sheet)s).Name.Value).ToList();
            else
                return new List<String>();
        }

        private List<String> GetExistingNamedRanges(Vt.VTVector vTVector_Two)
        {
            if (vTVector_Two.Count() > 0)
                return vTVector_Two.Elements<Vt.VTLPSTR>().Select(e => e.Text).ToList();
            else
                return new List<String>();
        }

        private List<String> GetModifiedNamedRanges()
        {
            var namedRanges = new List<String>();
            foreach (var w in Worksheets)
            {
                foreach (var n in w.NamedRanges)
                {
                    namedRanges.Add(w.Name + "!" + n.Name);
                }
                namedRanges.Add(w.Name + "!Print_Area");
                namedRanges.Add(w.Name + "!Print_Titles");
            }
            namedRanges.AddRange(NamedRanges.Select(n => n.Name));
            return namedRanges;
        }

        private void GenerateWorkbookPartContent(WorkbookPart workbookPart)
        {
            if (workbookPart.Workbook == null)
                workbookPart.Workbook = new Workbook();

            var workbook = workbookPart.Workbook;
            if (!workbook.NamespaceDeclarations.Contains(new KeyValuePair<string,string>("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")))
                workbook.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            #region WorkbookProperties
            if (workbook.WorkbookProperties == null)
                workbook.WorkbookProperties = new WorkbookProperties();

            if (workbook.WorkbookProperties.CodeName == null)
                workbook.WorkbookProperties.CodeName = "ThisWorkbook";

            if (workbook.WorkbookProperties.DefaultThemeVersion == null)
                workbook.WorkbookProperties.DefaultThemeVersion = (UInt32Value)124226U;
            #endregion

            if (workbook.BookViews == null)
                workbook.BookViews = new BookViews();

            if (workbook.Sheets == null)
                workbook.Sheets = new Sheets();

            var worksheets = (XLWorksheets)Worksheets;
            workbook.Sheets.Elements<Sheet>().Where(s => worksheets.Deleted.Contains(s.Id)).ForEach(s => s.Remove());

            foreach (var sheet in workbook.Sheets.Elements<Sheet>())
            {
                var sName = sheet.Name.Value;
                //if (Worksheets.Where(w => w.Name.ToLower() == sName.ToLower()).Any())
                if (Worksheets.Where(w=>(w as XLWorksheet).SheetId == (Int32)sheet.SheetId.Value).Any())
                {
                    var wks = (XLWorksheet)Worksheets.Where(w => (w as XLWorksheet).SheetId == (Int32)sheet.SheetId.Value).Single();
                    //wks.SheetId = (Int32)sheet.SheetId.Value;
                    wks.RelId = sheet.Id;
                    sheet.Name = wks.Name;
                }
            }

            foreach (var xlSheet in Worksheets.Cast<XLWorksheet>().Where(w => w.SheetId == 0).OrderBy(w => w.Position))
            {
                String rId = relId.GetNext(RelType.Workbook);
                while (Worksheets.Cast<XLWorksheet>().Where(w=>w.SheetId == Int32.Parse(rId.Substring(3))).Any())
                    rId = relId.GetNext(RelType.Workbook);

                xlSheet.SheetId = Int32.Parse(rId.Substring(3));
                xlSheet.RelId = rId;
                var newSheet = new Sheet()
                {
                    Name = xlSheet.Name,
                    Id = rId,
                    SheetId = (UInt32)xlSheet.SheetId
                };
                
                if (xlSheet.Visibility != XLWorksheetVisibility.Visible)
                    newSheet.State = sheetStateValues.Single(p => p.Key == xlSheet.Visibility).Value;

                workbook.Sheets.Append(newSheet);
            }

            var sheetElements = from sheet in workbook.Sheets.Elements<Sheet>()
                                join worksheet in Worksheets.Cast<XLWorksheet>() on sheet.Id.Value equals worksheet.RelId
                                orderby worksheet.Position
                                select sheet;

            UInt32 firstSheetVisible = 0;
            Boolean foundVisible = false;
            foreach (var sheet in sheetElements)
            {
                workbook.Sheets.RemoveChild(sheet);
                workbook.Sheets.Append(sheet);

                if (!foundVisible)
                    if (sheet.State == null || sheet.State == SheetStateValues.Visible)
                        foundVisible = true;
                    else
                        firstSheetVisible++;
            }

            WorkbookView workbookView = workbook.BookViews.Elements<WorkbookView>().FirstOrDefault();


            if (workbookView == null)
            {
                workbookView = new WorkbookView() { ActiveTab = firstSheetVisible, FirstSheet = firstSheetVisible };
                workbook.BookViews.Append(workbookView);
            }
            else
            {
                workbookView.ActiveTab = firstSheetVisible;
                workbookView.FirstSheet = firstSheetVisible;
            }
            

            DefinedNames definedNames = new DefinedNames();
            foreach (var worksheet in Worksheets.Cast<XLWorksheet>())
            {
                UInt32 sheetId = 0;
                foreach (var s in workbook.Sheets.Elements<Sheet>())
                {
                    if (s.SheetId == (UInt32)worksheet.SheetId)
                        break;
                    sheetId++;
                }

                if (worksheet.PageSetup.PrintAreas.Count() == 0)
                {
                    var minCell = worksheet.FirstCellUsed();
                    var maxCell = worksheet.LastCellUsed();
                    if (minCell != null && maxCell != null)
                        worksheet.PageSetup.PrintAreas.Add(minCell.Address.ToString(), maxCell.Address.ToString());
                }
                if (worksheet.PageSetup.PrintAreas.Count() > 0)
                {
                    DefinedName definedName = new DefinedName() { Name = "_xlnm.Print_Area", LocalSheetId = sheetId};
                    var definedNameText = String.Empty;
                    foreach (var printArea in worksheet.PageSetup.PrintAreas)
                    {
                        definedNameText += "'" + worksheet.Name + "'!"
                        + printArea.RangeAddress.FirstAddress.ToStringFixed()
                        + ":" + printArea.RangeAddress.LastAddress.ToStringFixed() + ",";
                    }
                    definedName.Text = definedNameText.Substring(0, definedNameText.Length - 1);
                    definedNames.Append(definedName);
                }

                foreach (var nr in worksheet.NamedRanges)
                {
                    DefinedName definedName = new DefinedName() { 
                        Name = nr.Name,
                        LocalSheetId = sheetId,
                        Text = nr.ToString()
                    };
                    if (!StringExtensions.IsNullOrWhiteSpace(nr.Comment)) definedName.Comment = nr.Comment;
                    definedNames.Append(definedName);
                }
                
                var titles = String.Empty;
                var definedNameTextRow = String.Empty;
                var definedNameTextColumn = String.Empty;
                if (worksheet.PageSetup.FirstRowToRepeatAtTop > 0)
                {
                    definedNameTextRow = "'" + worksheet.Name + "'!" + worksheet.PageSetup.FirstRowToRepeatAtTop.ToString()
                        + ":" + worksheet.PageSetup.LastRowToRepeatAtTop.ToString();
                }
                if (worksheet.PageSetup.FirstColumnToRepeatAtLeft > 0)
                {
                    var minColumn = worksheet.PageSetup.FirstColumnToRepeatAtLeft;
                    var maxColumn = worksheet.PageSetup.LastColumnToRepeatAtLeft;
                    definedNameTextColumn = "'" + worksheet.Name + "'!" + XLAddress.GetColumnLetterFromNumber(minColumn)
                        + ":" + XLAddress.GetColumnLetterFromNumber(maxColumn);
                }

                if (definedNameTextColumn.Length > 0)
                {
                    titles = definedNameTextColumn;
                    if (definedNameTextRow.Length > 0)
                        titles += "," + definedNameTextRow;
                }
                else
                {
                    titles = definedNameTextRow;
                }
                
                if (titles.Length > 0)
                {
                    DefinedName definedName = new DefinedName() { Name = "_xlnm.Print_Titles", LocalSheetId = sheetId };
                    definedName.Text = titles;
                    definedNames.Append(definedName);
                }
            }

            foreach (var nr in NamedRanges)
            {
                DefinedName definedName = new DefinedName()
                {
                    Name = nr.Name,
                    Text = nr.ToString()
                };
                if (!StringExtensions.IsNullOrWhiteSpace(nr.Comment)) definedName.Comment = nr.Comment;
                definedNames.Append(definedName);
            }

            if (workbook.DefinedNames == null)
                workbook.DefinedNames = new DefinedNames();

            foreach (DefinedName dn in definedNames)
            {
                if (workbook.DefinedNames.Elements<DefinedName>().Any(d =>
                d.Name.Value.ToLower() == dn.Name.Value.ToLower()
                    && (
                        (d.LocalSheetId != null && dn.LocalSheetId != null && d.LocalSheetId.InnerText == dn.LocalSheetId.InnerText)
                        || d.LocalSheetId == null || dn.LocalSheetId == null)
                    ))
                {
                    DefinedName existingDefinedName = (DefinedName)workbook.DefinedNames.Where(d => 
                        ((DefinedName)d).Name.Value.ToLower() == dn.Name.Value.ToLower()
                    && (
                        (((DefinedName)d).LocalSheetId != null && dn.LocalSheetId != null && ((DefinedName)d).LocalSheetId.InnerText == dn.LocalSheetId.InnerText)
                        || ((DefinedName)d).LocalSheetId == null || dn.LocalSheetId == null)
                        ).First();
                    existingDefinedName.Text = dn.Text;
                    existingDefinedName.LocalSheetId = dn.LocalSheetId;
                    existingDefinedName.Comment = dn.Comment;
                }
                else
                {
                    workbook.DefinedNames.Append(dn.CloneNode(true));
                }
            }

            if (workbook.CalculationProperties == null)
                workbook.CalculationProperties = new CalculationProperties() { CalculationId = (UInt32Value)125725U };

            if (CalculateMode == XLCalculateMode.Default)
                workbook.CalculationProperties.CalculationMode = null;
            else
                workbook.CalculationProperties.CalculationMode = calculateModeValues.Single(p => p.Key == CalculateMode).Value;


            if (ReferenceStyle == XLReferenceStyle.Default)
                workbook.CalculationProperties.ReferenceMode = null;
            else
                workbook.CalculationProperties.ReferenceMode = referenceModeValues.Single(p => p.Key == ReferenceStyle).Value;
           
        }

        private void GenerateSharedStringTablePartContent(SharedStringTablePart sharedStringTablePart)
        {
            HashSet<String> modifiedStrings = GetSharedStrings();

            sharedStringTablePart.SharedStringTable = new SharedStringTable() { Count = 0, UniqueCount = 0 };

            UInt32 stringCount = (UInt32)modifiedStrings.Count();

            Int32 stringId = 0;
            foreach (var s in modifiedStrings)
            {
                    SharedStringItem sharedStringItem = new SharedStringItem();
                    Text text = new Text();
                    text.Text = s;
                    sharedStringItem.Append(text);
                    sharedStringTablePart.SharedStringTable.Append(sharedStringItem);
                    sharedStringTablePart.SharedStringTable.Count += 1;
                    sharedStringTablePart.SharedStringTable.UniqueCount += 1;

                sharedStrings.Add(s, (UInt32)stringId);
                stringId++;
            }
        }

        #region GenerateWorkbookStylesPartContent
        private void GenerateWorkbookStylesPartContent(WorkbookStylesPart workbookStylesPart)
        {
            var defaultStyle = DefaultStyle;
            Dictionary<IXLFont, FontInfo> sharedFonts = new Dictionary<IXLFont, FontInfo>();
            sharedFonts.Add(defaultStyle.Font, new FontInfo() { FontId = 0, Font = defaultStyle.Font });

            Dictionary<IXLFill, FillInfo> sharedFills = new Dictionary<IXLFill, FillInfo>();
            sharedFills.Add(defaultStyle.Fill, new FillInfo() { FillId = 2, Fill = defaultStyle.Fill });

            Dictionary<IXLBorder, BorderInfo> sharedBorders = new Dictionary<IXLBorder, BorderInfo>();
            sharedBorders.Add(defaultStyle.Border, new BorderInfo() { BorderId = 0, Border = defaultStyle.Border });

            Dictionary<IXLNumberFormat, NumberFormatInfo> sharedNumberFormats = new Dictionary<IXLNumberFormat, NumberFormatInfo>();
            sharedNumberFormats.Add(defaultStyle.NumberFormat, new NumberFormatInfo() { NumberFormatId = 0, NumberFormat = defaultStyle.NumberFormat });

            //Dictionary<String, AlignmentInfo> sharedAlignments = new Dictionary<String, AlignmentInfo>();
            //sharedAlignments.Add(defaultStyle.Alignment.ToString(), new AlignmentInfo() { AlignmentId = 0, Alignment = defaultStyle.Alignment });

            if (workbookStylesPart.Stylesheet == null)
                workbookStylesPart.Stylesheet = new Stylesheet();

            // Cell styles = Named styles
            if (workbookStylesPart.Stylesheet.CellStyles == null)
                workbookStylesPart.Stylesheet.CellStyles = new CellStyles();

            UInt32 defaultFormatId;
            if (workbookStylesPart.Stylesheet.CellStyles.Elements<CellStyle>().Where(c => c.Name == "Normal").Any())
                defaultFormatId = workbookStylesPart.Stylesheet.CellStyles.Elements<CellStyle>().Where(c => c.Name == "Normal").Single().FormatId.Value;
            else
                if (workbookStylesPart.Stylesheet.CellStyles.Elements<CellStyle>().Any())
                    defaultFormatId = workbookStylesPart.Stylesheet.CellStyles.Elements<CellStyle>().Max(c => c.FormatId.Value) + 1;
                else
                    defaultFormatId = 0;

            sharedStyles.Add(defaultStyle,
                new StyleInfo()
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
            Int32 numberFormatCount = 1;
            var xlStyles = new List<IXLStyle>();


            foreach (var worksheet in Worksheets.Cast<XLWorksheet>())
            {
                xlStyles.AddRange(worksheet.Styles);
                worksheet.Internals.ColumnsCollection.Values.ForEach(c => xlStyles.Add(c.Style));
                worksheet.Internals.RowsCollection.Values.ForEach(c => xlStyles.Add(c.Style));
            }



            foreach (var xlStyle in xlStyles)
            {
                if (!sharedFonts.ContainsKey(xlStyle.Font))
                {
                    sharedFonts.Add(xlStyle.Font, new FontInfo() { FontId = fontCount++, Font = xlStyle.Font });
                }

                if (!sharedFills.ContainsKey(xlStyle.Fill))
                {
                    sharedFills.Add(xlStyle.Fill, new FillInfo() { FillId = fillCount++, Fill = xlStyle.Fill });
                }

                if (!sharedBorders.ContainsKey(xlStyle.Border))
                {
                    sharedBorders.Add(xlStyle.Border, new BorderInfo() { BorderId = borderCount++, Border = xlStyle.Border });
                }

                if (xlStyle.NumberFormat.NumberFormatId == -1 && !sharedNumberFormats.ContainsKey(xlStyle.NumberFormat))
                {
                    sharedNumberFormats.Add(xlStyle.NumberFormat, new NumberFormatInfo() { NumberFormatId = numberFormatCount + 164, NumberFormat = xlStyle.NumberFormat });
                    numberFormatCount++;
                }
            }
            

            var allSharedNumberFormats = ResolveNumberFormats(workbookStylesPart, sharedNumberFormats);
            var allSharedFonts = ResolveFonts(workbookStylesPart, sharedFonts);
            var allSharedFills = ResolveFills(workbookStylesPart, sharedFills);
            var allSharedBorders = ResolveBorders(workbookStylesPart, sharedBorders);

            foreach (var xlStyle in xlStyles)
            {
                if (!sharedStyles.ContainsKey(xlStyle))
                {
                    Int32 numberFormatId;
                    if (xlStyle.NumberFormat.NumberFormatId >= 0)
                        numberFormatId = xlStyle.NumberFormat.NumberFormatId;
                    else
                        numberFormatId = allSharedNumberFormats[xlStyle.NumberFormat].NumberFormatId;

                    sharedStyles.Add(xlStyle,
                        new StyleInfo()
                        {
                            StyleId = styleCount++,
                            Style = xlStyle,
                            FontId = allSharedFonts[xlStyle.Font].FontId,
                            FillId = allSharedFills[xlStyle.Fill].FillId,
                            BorderId = allSharedBorders[xlStyle.Border].BorderId,
                            NumberFormatId = numberFormatId
                        });
                }
            }

            var allCellStyleFormats = ResolveCellStyleFormats(workbookStylesPart);
            ResolveRest(workbookStylesPart);

            
            if (!workbookStylesPart.Stylesheet.CellStyles.Elements<CellStyle>().Where(c => c.Name == "Normal").Any())
            {
                //var defaultFormatId = sharedStyles.Values.Where(s => s.Style.Equals(DefaultStyle)).Single().StyleId;

                CellStyle cellStyle1 = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)defaultFormatId, BuiltinId = (UInt32Value)0U };
                workbookStylesPart.Stylesheet.CellStyles.Append(cellStyle1);
            }
            workbookStylesPart.Stylesheet.CellStyles.Count = (UInt32)workbookStylesPart.Stylesheet.CellStyles.Count();

            var newSharedStyles = new Dictionary<IXLStyle, StyleInfo>();
            foreach (var ss in sharedStyles)
            {
                Int32 styleId = -1;
                foreach (CellFormat f in workbookStylesPart.Stylesheet.CellFormats)
                {
                    styleId++;
                    if (CellFormatsAreEqual(f, ss.Value))
                        break;
                }
                if (styleId == -1) styleId = 0;
                var si = ss.Value;
                si.StyleId = (UInt32)styleId;
                newSharedStyles.Add(ss.Key, si);
            }
            sharedStyles.Clear();
            newSharedStyles.ForEach(kp => sharedStyles.Add(kp.Key, kp.Value));

            //TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium9", DefaultPivotStyle = "PivotStyleLight16" };
            //workbookStylesPart.Stylesheet.Append(tableStyles1);
        }

        private void ResolveRest(WorkbookStylesPart workbookStylesPart)
        {
            if (workbookStylesPart.Stylesheet.CellFormats == null)
                workbookStylesPart.Stylesheet.CellFormats = new CellFormats();

            foreach (var styleInfo in sharedStyles.Values)
            {
                Int32 styleId = 0;
                Boolean foundOne = false;
                foreach (CellFormat f in workbookStylesPart.Stylesheet.CellFormats)
                {
                    if (CellFormatsAreEqual(f, styleInfo))
                    {
                        foundOne = true;
                        break;
                    }
                    styleId++;
                }
                if (!foundOne)
                {
                    Int32 formatId = 0;
                    foreach (CellFormat f in workbookStylesPart.Stylesheet.CellStyleFormats)
                    {
                        if (CellFormatsAreEqual(f, styleInfo))
                            break;
                        styleId++;
                    }

                    //CellFormat cellFormat = new CellFormat() { NumberFormatId = (UInt32)styleInfo.NumberFormatId, FontId = (UInt32)styleInfo.FontId, FillId = (UInt32)styleInfo.FillId, BorderId = (UInt32)styleInfo.BorderId, ApplyNumberFormat = false, ApplyFill = ApplyFill(styleInfo), ApplyBorder = ApplyBorder(styleInfo), ApplyAlignment = false, ApplyProtection = false, FormatId = (UInt32)formatId };
                    CellFormat cellFormat = GetCellFormat(styleInfo);
                    cellFormat.FormatId = (UInt32)formatId;
                    Alignment alignment = new Alignment()
                    {
                        Horizontal = alignmentHorizontalValues.Single(a => a.Key == styleInfo.Style.Alignment.Horizontal).Value,
                        Vertical = alignmentVerticalValues.Single(a => a.Key == styleInfo.Style.Alignment.Vertical).Value,
                        Indent = (UInt32)styleInfo.Style.Alignment.Indent,
                        ReadingOrder = (UInt32)styleInfo.Style.Alignment.ReadingOrder,
                        WrapText = styleInfo.Style.Alignment.WrapText,
                        TextRotation = (UInt32)styleInfo.Style.Alignment.TextRotation,
                        ShrinkToFit = styleInfo.Style.Alignment.ShrinkToFit,
                        RelativeIndent = styleInfo.Style.Alignment.RelativeIndent,
                        JustifyLastLine = styleInfo.Style.Alignment.JustifyLastLine
                    };
                    cellFormat.Append(alignment);

                    if (cellFormat.ApplyProtection.Value)
                        cellFormat.Append(GetProtection(styleInfo));

                    workbookStylesPart.Stylesheet.CellFormats.Append(cellFormat);
                }
            }
            workbookStylesPart.Stylesheet.CellFormats.Count = (UInt32)workbookStylesPart.Stylesheet.CellFormats.Count();
        }

        private Dictionary<IXLStyle, StyleInfo> ResolveCellStyleFormats(WorkbookStylesPart workbookStylesPart)
        {
            if (workbookStylesPart.Stylesheet.CellStyleFormats == null)
                workbookStylesPart.Stylesheet.CellStyleFormats = new CellStyleFormats();

            var allSharedStyles = new Dictionary<IXLStyle, StyleInfo>();
            foreach (var styleInfo in sharedStyles.Values)
            {
                Int32 styleId = 0;
                Boolean foundOne = false;
                foreach (CellFormat f in workbookStylesPart.Stylesheet.CellStyleFormats)
                {
                    if (CellFormatsAreEqual(f, styleInfo))
                    {
                        foundOne = true;
                        break;
                    }
                    styleId++;
                }
                if (!foundOne)
                {
                    //CellFormat cellStyleFormat = new CellFormat() { NumberFormatId = (UInt32)styleInfo.NumberFormatId, FontId = (UInt32)styleInfo.FontId, FillId = (UInt32)styleInfo.FillId, BorderId = (UInt32)styleInfo.BorderId, ApplyNumberFormat = false, ApplyFill = ApplyFill(styleInfo), ApplyBorder = ApplyBorder(styleInfo), ApplyAlignment = false, ApplyProtection = false };
                    CellFormat cellStyleFormat = GetCellFormat(styleInfo);

                    if (cellStyleFormat.ApplyProtection.Value)
                        cellStyleFormat.Append(GetProtection(styleInfo));

                    workbookStylesPart.Stylesheet.CellStyleFormats.Append(cellStyleFormat);
                }
                allSharedStyles.Add(styleInfo.Style, new StyleInfo() { Style = styleInfo.Style, StyleId = (UInt32)styleId });
            }
            workbookStylesPart.Stylesheet.CellStyleFormats.Count = (UInt32)workbookStylesPart.Stylesheet.CellStyleFormats.Count();

            return allSharedStyles;
        }

        private Boolean ApplyFill(StyleInfo styleInfo)
        {
            return fillPatternValues.Single(p => p.Key == styleInfo.Style.Fill.PatternType).Value == PatternValues.None;
        }

        private Boolean ApplyBorder(StyleInfo styleInfo)
        {
            IXLBorder opBorder = styleInfo.Style.Border;
            return (
                   borderStyleValues.Single(b => b.Key == opBorder.BottomBorder).Value != BorderStyleValues.None
                || borderStyleValues.Single(b => b.Key == opBorder.DiagonalBorder).Value != BorderStyleValues.None
                || borderStyleValues.Single(b => b.Key == opBorder.RightBorder).Value != BorderStyleValues.None
                || borderStyleValues.Single(b => b.Key == opBorder.LeftBorder).Value != BorderStyleValues.None
                || borderStyleValues.Single(b => b.Key == opBorder.TopBorder).Value != BorderStyleValues.None);
        }

        private Boolean ApplyProtection(StyleInfo styleInfo)
        {
            return styleInfo.Style.Protection != null;
        }

        private CellFormat GetCellFormat(StyleInfo styleInfo)
        {
            var cellFormat = new CellFormat() { NumberFormatId = (UInt32)styleInfo.NumberFormatId, FontId = (UInt32)styleInfo.FontId, FillId = (UInt32)styleInfo.FillId, BorderId = (UInt32)styleInfo.BorderId, ApplyNumberFormat = false, ApplyFill = ApplyFill(styleInfo), ApplyBorder = ApplyBorder(styleInfo), ApplyAlignment = false, ApplyProtection = ApplyProtection(styleInfo) };
            return cellFormat;
        }

        private static Protection GetProtection(StyleInfo styleInfo)
        {
            return new Protection()
            {
                Locked = styleInfo.Style.Protection.Locked,
                Hidden = styleInfo.Style.Protection.Hidden
            };
        }

        private bool CellFormatsAreEqual(CellFormat f, StyleInfo styleInfo)
        {
            return
                   styleInfo.BorderId == f.BorderId
                && styleInfo.FillId == f.FillId
                && styleInfo.FontId == f.FontId
                && styleInfo.NumberFormatId == f.NumberFormatId
                && f.ApplyNumberFormat != null && f.ApplyNumberFormat == false
                && f.ApplyAlignment != null && f.ApplyAlignment == false
                && f.ApplyFill != null && f.ApplyFill == ApplyFill(styleInfo)
                && f.ApplyBorder != null && f.ApplyBorder == ApplyBorder(styleInfo)
                && AlignmentsAreEqual(f.Alignment, styleInfo.Style.Alignment)
                && ProtectionsAreEqual(f.Protection, styleInfo.Style.Protection)
                ;
        }

        private bool ProtectionsAreEqual(Protection protection, IXLProtection xlProtection)
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


        private bool AlignmentsAreEqual(Alignment alignment, IXLAlignment xlAlignment)
        {
            var a = new XLAlignment();
            if (alignment != null)
            {
                if (alignment.Horizontal != null)
                    a.Horizontal = alignmentHorizontalValues.Single(p => p.Value == alignment.Horizontal.Value).Key;
                if (alignment.Vertical != null)
                    a.Vertical = alignmentVerticalValues.Single(p => p.Value == alignment.Vertical.Value).Key;
                if (alignment.Indent != null)
                    a.Indent = (Int32)alignment.Indent.Value;
                if (alignment.ReadingOrder != null)
                    a.ReadingOrder = alignmentReadingOrderValues.Single(p => p.Value == alignment.ReadingOrder.Value).Key;
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

        private Dictionary<IXLBorder, BorderInfo> ResolveBorders(WorkbookStylesPart workbookStylesPart, Dictionary<IXLBorder, BorderInfo> sharedBorders)
        {
            if (workbookStylesPart.Stylesheet.Borders == null)
                workbookStylesPart.Stylesheet.Borders = new Borders();

            var allSharedBorders = new Dictionary<IXLBorder, BorderInfo>();
            foreach (var borderInfo in sharedBorders.Values)
            {
                Int32 borderId = 0;
                Boolean foundOne = false;
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
                    Border border = GetNewBorder(borderInfo);
                    workbookStylesPart.Stylesheet.Borders.Append(border);
                }
                allSharedBorders.Add(borderInfo.Border, new BorderInfo() { Border = borderInfo.Border, BorderId = (UInt32)borderId });
            }
            workbookStylesPart.Stylesheet.Borders.Count = (UInt32)workbookStylesPart.Stylesheet.Borders.Count();
            return allSharedBorders;
        }

        private Border GetNewBorder(BorderInfo borderInfo)
        {
            Border border = new Border() { DiagonalUp = borderInfo.Border.DiagonalUp, DiagonalDown = borderInfo.Border.DiagonalDown };

            LeftBorder leftBorder = new LeftBorder() { Style = borderStyleValues.Single(b => b.Key == borderInfo.Border.LeftBorder).Value };
            Color leftBorderColor = GetNewColor(borderInfo.Border.LeftBorderColor);
            leftBorder.Append(leftBorderColor);
            border.Append(leftBorder);

            RightBorder rightBorder = new RightBorder() { Style = borderStyleValues.Single(b => b.Key == borderInfo.Border.RightBorder).Value };
            Color rightBorderColor = GetNewColor(borderInfo.Border.RightBorderColor);
            rightBorder.Append(rightBorderColor);
            border.Append(rightBorder);

            TopBorder topBorder = new TopBorder() { Style = borderStyleValues.Single(b => b.Key == borderInfo.Border.TopBorder).Value };
            Color topBorderColor = GetNewColor(borderInfo.Border.TopBorderColor);
            topBorder.Append(topBorderColor);
            border.Append(topBorder);

            BottomBorder bottomBorder = new BottomBorder() { Style = borderStyleValues.Single(b => b.Key == borderInfo.Border.BottomBorder).Value };
            Color bottomBorderColor = GetNewColor(borderInfo.Border.BottomBorderColor);
            bottomBorder.Append(bottomBorderColor);
            border.Append(bottomBorder);

            DiagonalBorder diagonalBorder = new DiagonalBorder() { Style = borderStyleValues.Single(b => b.Key == borderInfo.Border.DiagonalBorder).Value };
            Color diagonalBorderColor = GetNewColor(borderInfo.Border.DiagonalBorderColor);
            diagonalBorder.Append(diagonalBorderColor);
            border.Append(diagonalBorder);

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
                    nb.LeftBorder = borderStyleValues.Single(p => p.Value == b.LeftBorder.Style).Key;
                var bColor = GetColor(b.LeftBorder.Color);
                if (bColor.HasValue)
                    nb.LeftBorderColor = bColor;
            }

            if (b.RightBorder != null)
            {
                if (b.RightBorder.Style != null)
                    nb.RightBorder = borderStyleValues.Single(p => p.Value == b.RightBorder.Style).Key;
                var bColor = GetColor(b.RightBorder.Color);
                if (bColor.HasValue)
                    nb.RightBorderColor = bColor;
            }

            if (b.TopBorder != null)
            {
                if (b.TopBorder.Style != null)
                    nb.TopBorder = borderStyleValues.Single(p => p.Value == b.TopBorder.Style).Key;
                var bColor = GetColor(b.TopBorder.Color);
                if (bColor.HasValue)
                    nb.TopBorderColor = bColor;
            }

            if (b.BottomBorder != null)
            {
                if (b.BottomBorder.Style != null)
                    nb.BottomBorder = borderStyleValues.Single(p => p.Value == b.BottomBorder.Style).Key;
                var bColor = GetColor(b.BottomBorder.Color);
                if (bColor.HasValue)
                    nb.BottomBorderColor = bColor;
            }

            return nb.Equals(xlBorder);
        }

        private Dictionary<IXLFill, FillInfo> ResolveFills(WorkbookStylesPart workbookStylesPart, Dictionary<IXLFill, FillInfo> sharedFills)
        {
            if (workbookStylesPart.Stylesheet.Fills == null)
                workbookStylesPart.Stylesheet.Fills = new Fills();

            ResolveFillWithPattern(workbookStylesPart.Stylesheet.Fills, PatternValues.None);
            ResolveFillWithPattern(workbookStylesPart.Stylesheet.Fills, PatternValues.Gray125);

            var allSharedFills = new Dictionary<IXLFill, FillInfo>();
            foreach (var fillInfo in sharedFills.Values)
            {
                Int32 fillId = 0;
                Boolean foundOne = false;
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
                    Fill fill = GetNewFill(fillInfo);
                    workbookStylesPart.Stylesheet.Fills.Append(fill);
                }
                allSharedFills.Add(fillInfo.Fill, new FillInfo() { Fill = fillInfo.Fill, FillId = (UInt32)fillId });
            }
         
            workbookStylesPart.Stylesheet.Fills.Count = (UInt32)workbookStylesPart.Stylesheet.Fills.Count();
            return allSharedFills;
        }

        private void ResolveFillWithPattern(Fills fills, PatternValues patternValues)
        {
            if (!fills.Elements<Fill>().Where(f => 
                f.PatternFill.PatternType == patternValues
                && f.PatternFill.ForegroundColor == null
                && f.PatternFill.BackgroundColor == null
                ).Any())
            {
                Fill fill1 = new Fill();
                PatternFill patternFill1 = new PatternFill() { PatternType = patternValues };
                fill1.Append(patternFill1);
                fills.Append(fill1);
            }
            
        }

        private Fill GetNewFill(FillInfo fillInfo)
        {
            Fill fill = new Fill();

            PatternFill patternFill = new PatternFill() { PatternType = fillPatternValues.Single(p => p.Key == fillInfo.Fill.PatternType).Value };
            ForegroundColor foregroundColor = new ForegroundColor();
            if (fillInfo.Fill.PatternColor.ColorType == XLColorType.Color)
                foregroundColor.Rgb = fillInfo.Fill.PatternColor.Color.ToHex();
            else if (fillInfo.Fill.PatternColor.ColorType == XLColorType.Indexed)
                foregroundColor.Indexed = (UInt32)fillInfo.Fill.PatternColor.Indexed;
            else
            {
                foregroundColor.Theme = (UInt32)fillInfo.Fill.PatternColor.ThemeColor;
                if (fillInfo.Fill.PatternColor.ThemeTint != 1)
                    foregroundColor.Tint = fillInfo.Fill.PatternColor.ThemeTint;
            }
            BackgroundColor backgroundColor = new BackgroundColor();
            if (fillInfo.Fill.PatternBackgroundColor.ColorType == XLColorType.Color)
                backgroundColor.Rgb = fillInfo.Fill.PatternBackgroundColor.Color.ToHex();
            else if (fillInfo.Fill.PatternBackgroundColor.ColorType == XLColorType.Indexed)
                backgroundColor.Indexed = (UInt32)fillInfo.Fill.PatternBackgroundColor.Indexed;
            else
            {
                backgroundColor.Theme = (UInt32)fillInfo.Fill.PatternBackgroundColor.ThemeColor;
                if (fillInfo.Fill.PatternBackgroundColor.ThemeTint != 1)
                    backgroundColor.Tint = fillInfo.Fill.PatternBackgroundColor.ThemeTint;
            }

            patternFill.Append(foregroundColor);
            patternFill.Append(backgroundColor);

            fill.Append(patternFill);

            return fill;
        }

        private bool FillsAreEqual(Fill f, IXLFill xlFill)
        {
            var nF = new XLFill();
            if (f.PatternFill != null)
            {
                if (f.PatternFill.PatternType != null)
                    nF.PatternType = fillPatternValues.Single(p => p.Value == f.PatternFill.PatternType).Key;

                var fColor = GetColor(f.PatternFill.ForegroundColor);
                if (fColor.HasValue)
                    nF.PatternColor = fColor;

                var bColor = GetColor(f.PatternFill.BackgroundColor);
                if (bColor.HasValue)
                    nF.PatternBackgroundColor = bColor;
            }
            return nF.Equals(xlFill);
        }

        private Dictionary<IXLFont, FontInfo> ResolveFonts(WorkbookStylesPart workbookStylesPart, Dictionary<IXLFont, FontInfo> sharedFonts)
        {
            if (workbookStylesPart.Stylesheet.Fonts == null)
                workbookStylesPart.Stylesheet.Fonts = new Fonts();

            var allSharedFonts = new Dictionary<IXLFont, FontInfo>();
            foreach (var fontInfo in sharedFonts.Values)
            {
                Int32 fontId = 0;
                Boolean foundOne = false;
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
                    Font font = GetNewFont(fontInfo);
                    workbookStylesPart.Stylesheet.Fonts.Append(font);
                }
                allSharedFonts.Add(fontInfo.Font, new FontInfo() { Font = fontInfo.Font, FontId = (UInt32)fontId });
            }
            workbookStylesPart.Stylesheet.Fonts.Count = (UInt32)workbookStylesPart.Stylesheet.Fonts.Count();
            return allSharedFonts;
        }

        private Font GetNewFont(FontInfo fontInfo)
        {
            Font font = new Font();
            Bold bold = fontInfo.Font.Bold ? new Bold() : null;
            Italic italic = fontInfo.Font.Italic ? new Italic() : null;
            Underline underline = fontInfo.Font.Underline != XLFontUnderlineValues.None ? new Underline() { Val = underlineValuesList.Single(u => u.Key == fontInfo.Font.Underline).Value } : null;
            Strike strike = fontInfo.Font.Strikethrough ? new Strike() : null;
            VerticalTextAlignment verticalAlignment = new VerticalTextAlignment() { Val = fontVerticalTextAlignmentValues.Single(f => f.Key == fontInfo.Font.VerticalAlignment).Value };
            Shadow shadow = fontInfo.Font.Shadow ? new Shadow() : null;
            FontSize fontSize = new FontSize() { Val = fontInfo.Font.FontSize };
            Color color = GetNewColor(fontInfo.Font.FontColor);
            
            FontName fontName = new FontName() { Val = fontInfo.Font.FontName };
            FontFamilyNumbering fontFamilyNumbering = new FontFamilyNumbering() { Val = (Int32)fontInfo.Font.FontFamilyNumbering };

            if (bold != null) font.Append(bold);
            if (italic != null) font.Append(italic);
            if (underline != null) font.Append(underline);
            if (strike != null) font.Append(strike);
            font.Append(verticalAlignment);
            if (shadow != null) font.Append(shadow);
            font.Append(fontSize);
            font.Append(color);
            font.Append(fontName);
            font.Append(fontFamilyNumbering);

            return font;
        }

        private Color GetNewColor(IXLColor xlColor)
        {
            Color color = new Color();
            if (xlColor.ColorType == XLColorType.Color)
                color.Rgb = xlColor.Color.ToHex();
            else if (xlColor.ColorType == XLColorType.Indexed)
                color.Indexed = (UInt32)xlColor.Indexed;
            else
            {
                color.Theme = (UInt32)xlColor.ThemeColor;
                if (xlColor.ThemeTint != 1)
                    color.Tint = xlColor.ThemeTint;
            }
            return color;
        }

        private bool FontsAreEqual(Font f, IXLFont xlFont)
        {
            var nf = new XLFont();
            nf.Bold = f.Bold != null;
            nf.Italic = f.Italic != null;
            if (f.Underline != null)
                if (f.Underline.Val != null)
                    nf.Underline = underlineValuesList.Single(u => u.Value == f.Underline.Val).Key;
                else
                    nf.Underline = XLFontUnderlineValues.Single;
            nf.Strikethrough = f.Strike != null;
            if (f.VerticalTextAlignment != null)
                if (f.VerticalTextAlignment.Val != null)
                    nf.VerticalAlignment = fontVerticalTextAlignmentValues.Single(v => v.Value == f.VerticalTextAlignment.Val).Key;
                else
                    nf.VerticalAlignment = XLFontVerticalTextAlignmentValues.Baseline;
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

        private Dictionary<IXLNumberFormat, NumberFormatInfo> ResolveNumberFormats(WorkbookStylesPart workbookStylesPart, Dictionary<IXLNumberFormat, NumberFormatInfo> sharedNumberFormats)
        {
            if (workbookStylesPart.Stylesheet.NumberingFormats == null)
                workbookStylesPart.Stylesheet.NumberingFormats = new NumberingFormats();

            var allSharedNumberFormats = new Dictionary<IXLNumberFormat, NumberFormatInfo>();
            foreach (var numberFormatInfo in sharedNumberFormats.Values)
            {
                Int32 numberingFormatId = 0;
                Boolean foundOne = false;
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
                    NumberingFormat numberingFormat = new NumberingFormat() { NumberFormatId = (UInt32)numberingFormatId, FormatCode = numberFormatInfo.NumberFormat.Format };
                    workbookStylesPart.Stylesheet.NumberingFormats.Append(numberingFormat);
                }
                allSharedNumberFormats.Add(numberFormatInfo.NumberFormat, new NumberFormatInfo() { NumberFormat = numberFormatInfo.NumberFormat, NumberFormatId = numberingFormatId });
            }
            workbookStylesPart.Stylesheet.NumberingFormats.Count = (UInt32)workbookStylesPart.Stylesheet.NumberingFormats.Count();
            return allSharedNumberFormats;
        }

        private bool NumberFormatsAreEqual(NumberingFormat nf, IXLNumberFormat xlNumberFormat)
        {
            var newXLNumberFormat = new XLNumberFormat();
            
            if (nf.FormatCode != null && !StringExtensions.IsNullOrWhiteSpace(nf.FormatCode.Value))
                newXLNumberFormat.Format = nf.FormatCode.Value;
            else if (nf.NumberFormatId != null)
                newXLNumberFormat.NumberFormatId = (Int32)nf.NumberFormatId.Value;

            return newXLNumberFormat.Equals(xlNumberFormat);
        }
#endregion

        #region GenerateWorksheetPartContent
        private void GenerateWorksheetPartContent(WorksheetPart worksheetPart, XLWorksheet xlWorksheet)
        {
            
            #region Worksheet
            if (worksheetPart.Worksheet == null)
                worksheetPart.Worksheet = new Worksheet();

            GenerateTables(xlWorksheet, worksheetPart);

            if (!worksheetPart.Worksheet.NamespaceDeclarations.Contains(new KeyValuePair<String, String>("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")))
                worksheetPart.Worksheet.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            #endregion

            var cm = new XLWSContentManager(worksheetPart.Worksheet);

            #region SheetProperties
            if (worksheetPart.Worksheet.SheetProperties == null)
                worksheetPart.Worksheet.SheetProperties = new SheetProperties() { CodeName = xlWorksheet.Name.RemoveSpecialCharacters() };

            cm.SetElement(XLWSContentManager.XLWSContents.SheetProperties, worksheetPart.Worksheet.SheetProperties);

            if (worksheetPart.Worksheet.SheetProperties.OutlineProperties == null)
                worksheetPart.Worksheet.SheetProperties.OutlineProperties = new OutlineProperties();

            worksheetPart.Worksheet.SheetProperties.OutlineProperties.SummaryBelow = (xlWorksheet.Outline.SummaryVLocation == XLOutlineSummaryVLocation.Bottom);
            worksheetPart.Worksheet.SheetProperties.OutlineProperties.SummaryRight = (xlWorksheet.Outline.SummaryHLocation == XLOutlineSummaryHLocation.Right);

            if (worksheetPart.Worksheet.SheetProperties.PageSetupProperties == null && (xlWorksheet.PageSetup.PagesTall > 0 || xlWorksheet.PageSetup.PagesWide > 0))
                worksheetPart.Worksheet.SheetProperties.PageSetupProperties = new PageSetupProperties();

            if (xlWorksheet.PageSetup.PagesTall > 0 || xlWorksheet.PageSetup.PagesWide > 0)
                worksheetPart.Worksheet.SheetProperties.PageSetupProperties.FitToPage = true;

            #endregion

            UInt32 maxColumn = 0;
            UInt32 maxRow = 0;

            String sheetDimensionReference = "A1";
            if (xlWorksheet.Internals.CellsCollection.Count > 0)
            {
                maxColumn = (UInt32)xlWorksheet.Internals.CellsCollection.Select(c => c.Key.ColumnNumber).Max();
                maxRow = (UInt32)xlWorksheet.Internals.CellsCollection.Select(c => c.Key.RowNumber).Max();
                sheetDimensionReference = "A1:" + XLAddress.GetColumnLetterFromNumber((Int32)maxColumn) + ((Int32)maxRow).ToStringLookup();
            }

            if (xlWorksheet.Internals.ColumnsCollection.Count > 0)
            {
                UInt32 maxColCollection = (UInt32)xlWorksheet.Internals.ColumnsCollection.Keys.Max();
                if (maxColCollection > maxColumn) maxColumn = maxColCollection;
            }

            if (xlWorksheet.Internals.RowsCollection.Count > 0)
            {
                UInt32 maxRowCollection = (UInt32)xlWorksheet.Internals.RowsCollection.Keys.Max();
                if (maxRowCollection > maxRow) maxRow = maxRowCollection;
            }

            #region SheetViews
            if (worksheetPart.Worksheet.SheetDimension == null)
                worksheetPart.Worksheet.SheetDimension = new SheetDimension() { Reference = sheetDimensionReference };

            cm.SetElement(XLWSContentManager.XLWSContents.SheetDimension, worksheetPart.Worksheet.SheetDimension);

            if (worksheetPart.Worksheet.SheetViews == null)
                worksheetPart.Worksheet.SheetViews = new SheetViews();

            cm.SetElement(XLWSContentManager.XLWSContents.SheetViews, worksheetPart.Worksheet.SheetViews);

            var sheetView = worksheetPart.Worksheet.SheetViews.FirstOrDefault();
            if (worksheetPart.Worksheet.SheetViews.Count() == 0)
            {
                sheetView = new SheetView() { WorkbookViewId = (UInt32Value)0U };
                worksheetPart.Worksheet.SheetViews.Append(sheetView);
            }


            var pane = sheetView.Elements<Pane>().FirstOrDefault();
            if (pane == null)
            {
                pane = new Pane();
                sheetView.Append(pane);
            }

            Double hSplit = 0;
            Double ySplit = 0;
            //if (xlWorksheet.SheetView.FreezePanes)
            //{
            pane.State = PaneStateValues.FrozenSplit;
            hSplit = xlWorksheet.SheetView.SplitColumn;
            ySplit = xlWorksheet.SheetView.SplitRow;
            //}
            //else
            //{
            //    pane.State = null;
            //    foreach (var column in xlWorksheet.Columns(1, xlWorksheet.SheetView.SplitColumn))
            //    {
            //        hSplit += (column.Width * 141.33);
            //    }
            //    foreach (var row in xlWorksheet.Rows(1, xlWorksheet.SheetView.SplitRow))
            //    {
            //        ySplit += (row.Height * 37.0);
            //    }
            //}

            pane.HorizontalSplit = hSplit;
            pane.VerticalSplit = ySplit;

            pane.TopLeftCell = XLAddress.GetColumnLetterFromNumber(xlWorksheet.SheetView.SplitColumn + 1)
                                    + (xlWorksheet.SheetView.SplitRow + 1).ToString();

            if (hSplit == 0 && ySplit == 0)
                sheetView.RemoveAllChildren<Pane>();

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

            cm.SetElement(XLWSContentManager.XLWSContents.SheetFormatProperties, worksheetPart.Worksheet.SheetFormatProperties);

            worksheetPart.Worksheet.SheetFormatProperties.DefaultRowHeight = xlWorksheet.RowHeight;
            worksheetPart.Worksheet.SheetFormatProperties.DefaultColumnWidth = xlWorksheet.ColumnWidth;
            if (xlWorksheet.RowHeightChanged)
                worksheetPart.Worksheet.SheetFormatProperties.CustomHeight = true;
            

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
            Columns columns = null;
            if (xlWorksheet.Internals.CellsCollection.Count == 0 && xlWorksheet.Internals.ColumnsCollection.Count == 0)
            {
                worksheetPart.Worksheet.RemoveAllChildren<Columns>();
            }
            else
            {
                var worksheetColumnWidth = GetColumnWidth(xlWorksheet.ColumnWidth);

                if (worksheetPart.Worksheet.Elements<Columns>().Count() == 0)
                {
                    var previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.Columns);
                    worksheetPart.Worksheet.InsertAfter(new Columns(), previousElement);
                }

                columns = worksheetPart.Worksheet.Elements<Columns>().First();
                cm.SetElement(XLWSContentManager.XLWSContents.Columns, columns);

                Dictionary<UInt32, Column> sheetColumnsByMin = columns.Elements<Column>().ToDictionary(c => c.Min.Value, c => c);
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

                if (minInColumnsCollection > 1)
                {
                    UInt32Value min = 1;
                    UInt32Value max = (UInt32)(minInColumnsCollection - 1);
                    var styleId = sharedStyles[xlWorksheet.Style].StyleId;

                    for (var co = min; co <= max; co++)
                    {
                        Column column = new Column()
                        {
                            Min = co,
                            Max = co,
                            Style = styleId,
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
                    Boolean isHidden = false;
                    Boolean collapsed = false;
                    Int32 outlineLevel = 0;
                    if (xlWorksheet.Internals.ColumnsCollection.ContainsKey(co))
                    {
                        styleId = sharedStyles[xlWorksheet.Internals.ColumnsCollection[co].Style].StyleId;
                        columnWidth = GetColumnWidth(xlWorksheet.Internals.ColumnsCollection[co].Width);
                        isHidden = xlWorksheet.Internals.ColumnsCollection[co].IsHidden;
                        collapsed = xlWorksheet.Internals.ColumnsCollection[co].Collapsed;
                        outlineLevel = xlWorksheet.Internals.ColumnsCollection[co].OutlineLevel;
                    }
                    else
                    {
                        styleId = sharedStyles[xlWorksheet.Style].StyleId;
                        columnWidth = worksheetColumnWidth;
                    }

                    Column column = new Column()
                    {
                        Min = (UInt32)co,
                        Max = (UInt32)co,
                        Style = styleId,
                        Width = columnWidth,
                        CustomWidth = true
                    };
                    if (isHidden) column.Hidden = true;
                    if (collapsed) column.Collapsed = true;
                    if (outlineLevel > 0) column.OutlineLevel = (byte)outlineLevel;

                    UpdateColumn(column, columns, sheetColumnsByMin); //, sheetColumnsByMax);
                }

                foreach (var col in columns.Elements<Column>().Where(c => c.Min > (UInt32)(maxInColumnsCollection)).OrderBy(c => c.Min.Value))
                {
                    col.Style = sharedStyles[xlWorksheet.Style].StyleId;
                    col.Width = worksheetColumnWidth;
                    col.CustomWidth = true;
                    if ((Int32)col.Max.Value > maxInColumnsCollection)
                        maxInColumnsCollection = (Int32)col.Max.Value;
                }

                if (maxInColumnsCollection < XLWorksheet.MaxNumberOfColumns)
                {
                    Column column = new Column()
                    {
                        Min = (UInt32)(maxInColumnsCollection + 1),
                        Max = (UInt32)(XLWorksheet.MaxNumberOfColumns),
                        Style = sharedStyles[xlWorksheet.Style].StyleId,
                        Width = worksheetColumnWidth,
                        CustomWidth = true
                    };
                    columns.Append(column);
                }

                CollapseColumns(columns, sheetColumnsByMin);
            }
            #endregion

            #region SheetData
            SheetData sheetData;
            if (worksheetPart.Worksheet.Elements<SheetData>().Count() == 0)
            {
                OpenXmlElement previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.SheetData);
                worksheetPart.Worksheet.InsertAfter(new SheetData(), previousElement);
            }

            sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
            cm.SetElement(XLWSContentManager.XLWSContents.SheetData, sheetData);

            var cellsByRow = new Dictionary<Int32, List<IXLCell>>();
            foreach (var c in xlWorksheet.Internals.CellsCollection.Values)
            {
                Int32 rowNum = c.Address.RowNumber;
                if (!cellsByRow.ContainsKey(rowNum))
                    cellsByRow.Add(rowNum, new List<IXLCell>());

                cellsByRow[rowNum].Add(c);
            }

            var distinctRows = cellsByRow.Keys.Union(xlWorksheet.Internals.RowsCollection.Keys);
            Boolean noRows = (sheetData.Elements<Row>().FirstOrDefault() == null);
            var sheetDataRows = sheetData.Elements<Row>().ToDictionary(r => (Int32)r.RowIndex.Value, r => r);
            foreach (var distinctRow in distinctRows.OrderBy(r => r))
            {
                Row row; // = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex.Value == (UInt32)distinctRow);
                if (sheetDataRows.ContainsKey(distinctRow))
                {
                    row = sheetDataRows[distinctRow];
                }
                else
                {
                    row = new Row() { RowIndex = (UInt32)distinctRow };
                    if (noRows)
                    {
                        sheetData.Append(row);
                        noRows = false;
                    }
                    else
                    {
                        if (sheetDataRows.Where(r => r.Key > row.RowIndex.Value).Any())
                        {
                            var minRow = sheetDataRows.Where(r => r.Key > (Int32)row.RowIndex.Value).Min(r=>r.Key);
                            Row rowBeforeInsert = sheetDataRows[minRow];
                            sheetData.InsertBefore(row, rowBeforeInsert);
                        }
                        else
                        {
                            sheetData.Append(row);
                        }
                    }
                }

                if (maxColumn > 0)
                    row.Spans = new ListValue<StringValue>() { InnerText = "1:" + maxColumn.ToString() };

                row.Height = null;
                row.CustomHeight = null;
                row.Hidden = null;
                row.StyleIndex = null;
                row.CustomFormat = null;
                if (xlWorksheet.Internals.RowsCollection.ContainsKey(distinctRow))
                {
                    var thisRow = xlWorksheet.Internals.RowsCollection[distinctRow];
                    if (thisRow.Height != xlWorksheet.RowHeight)
                    {
                        row.Height = thisRow.Height;
                        row.CustomHeight = true;
                    }
                    if (!thisRow.Style.Equals(xlWorksheet.Style))
                    {
                        row.StyleIndex = sharedStyles[thisRow.Style].StyleId;
                        row.CustomFormat = true;
                    }
                    if (thisRow.IsHidden) row.Hidden = true;
                    if (thisRow.Collapsed) row.Collapsed = true;
                    if (thisRow.OutlineLevel > 0) row.OutlineLevel = (byte)thisRow.OutlineLevel;
                }
                else
                {
                    //row.Height = xlWorksheet.RowHeight;
                    //row.CustomHeight = true;
                    //row.Hidden = false;
                }

                List<Cell> cellsToRemove = new List<Cell>();
                foreach (var cell in row.Elements<Cell>())
                {
                    var cellReference = cell.CellReference;
                    if (xlWorksheet.Internals.CellsCollection.Deleted.ContainsKey(new XLAddress(cellReference)))
                        cellsToRemove.Add(cell);
                }
                cellsToRemove.ForEach(cell => row.RemoveChild(cell));

                
                if (cellsByRow.ContainsKey(distinctRow))
                {

                    foreach (var opCell in cellsByRow[distinctRow]
                        .OrderBy(c => c.Address.ColumnNumber)
                        .Select(c => (XLCell)c))
                    {
                        var styleId = sharedStyles[opCell.Style].StyleId;

                        var dataType = opCell.DataType;
                        var cellReference = ((XLAddress)opCell.Address).GetTrimmedAddress();

                        //Boolean isNewCell = false;
                        
                        Cell cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference.Value == cellReference);
                        if (cell == null)
                        {
                            //isNewCell = true;
                            cell = new Cell() { CellReference = cellReference };
                            if (row.Elements<Cell>().Count() == 0)
                            {
                                row.Append(cell);
                            }
                            else
                            {
                                Int32 newColumn = XLAddress.GetColumnNumberFromAddress1(cellReference);

                                Cell cellBeforeInsert = null;
                                Int32 lastCo = Int32.MaxValue;
                                foreach (var c in row.Elements<Cell>().Where(c => XLAddress.GetColumnNumberFromAddress1(c.CellReference.Value) > newColumn))
                                {
                                    var thidCo = XLAddress.GetColumnNumberFromAddress1(c.CellReference.Value);
                                    if (lastCo > thidCo)
                                    {
                                        cellBeforeInsert = c;
                                        lastCo = thidCo;
                                    }
                                }
                                if (cellBeforeInsert == null)
                                    row.Append(cell);
                                else
                                    row.InsertBefore(cell, cellBeforeInsert);
                            }
                        }

                        cell.StyleIndex = styleId;
                        if (!StringExtensions.IsNullOrWhiteSpace(opCell.FormulaA1))
                        {
                            String formula = opCell.FormulaA1;
                            if (formula.StartsWith("{"))
                            {
                                formula = formula.Substring(1, formula.Length - 2);
                                cell.CellFormula = new CellFormula(formula);
                                cell.CellFormula.FormulaType = CellFormulaValues.Array;
                                cell.CellFormula.Reference = cellReference;
                            }
                            else
                            {
                                cell.CellFormula = new CellFormula(formula);
                            }
                            cell.CellValue = null;
                        }
                        else
                        {
                            cell.CellFormula = null;

                            if (opCell.DataType == XLCellValues.DateTime)
                                cell.DataType = null;
                            else
                                cell.DataType = GetCellValue(opCell);
                            
                            CellValue cellValue = new CellValue();
                            if (dataType == XLCellValues.Text)
                            {
                                if (StringExtensions.IsNullOrWhiteSpace(opCell.InnerText))
                                {
                                    cell.CellValue = null;
                                }
                                else
                                {
                                    if (opCell.ShareString)
                                    {
                                        cellValue.Text = sharedStrings[opCell.InnerText].ToString();
                                        cell.CellValue = cellValue;
                                    }
                                    else
                                    {
                                        cell.InlineString = new InlineString() { Text = new Text(opCell.GetString()) };
                                    }
                                }
                            }
                            else if (dataType == XLCellValues.TimeSpan)
                            {
                                TimeSpan timeSpan = opCell.GetTimeSpan();
                                cellValue.Text = XLCell.baseDate.Add(timeSpan).ToOADate().ToString(CultureInfo.InvariantCulture);
                                cell.CellValue = cellValue;
                            }
                            else if (dataType == XLCellValues.DateTime || dataType == XLCellValues.Number)
                            {
                                cellValue.Text = Double.Parse(opCell.InnerText).ToString(CultureInfo.InvariantCulture);
                                cell.CellValue = cellValue;
                            }
                            else
                            {
                                cellValue.Text = opCell.InnerText;
                                cell.CellValue = cellValue;
                            }
                        }
                    }
                }
            }
            #endregion

            #region SheetProtection
            SheetProtection sheetProtection = null;
            if (xlWorksheet.Protection.Protected)
            {
                if (worksheetPart.Worksheet.Elements<SheetProtection>().Count() == 0)
                {
                    OpenXmlElement previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.SheetProtection);
                    worksheetPart.Worksheet.InsertAfter(new SheetProtection(), previousElement);
                }

                sheetProtection = worksheetPart.Worksheet.Elements<SheetProtection>().First();
                cm.SetElement(XLWSContentManager.XLWSContents.SheetProtection, sheetProtection);

                var protection = (XLSheetProtection)xlWorksheet.Protection;
                sheetProtection.Sheet = protection.Protected;
                if (!StringExtensions.IsNullOrWhiteSpace(protection.PasswordHash))
                    sheetProtection.Password = protection.PasswordHash;
                sheetProtection.FormatCells = protection.FormatCells;
                sheetProtection.FormatColumns = protection.FormatColumns;
                sheetProtection.FormatRows = protection.FormatRows;
                sheetProtection.InsertColumns = protection.InsertColumns;
                sheetProtection.InsertHyperlinks = protection.InsertHyperlinks;
                sheetProtection.InsertRows = protection.InsertRows;
                sheetProtection.DeleteColumns = protection.DeleteColumns;
                sheetProtection.DeleteRows = protection.DeleteRows;
                sheetProtection.AutoFilter = protection.AutoFilter;
                sheetProtection.PivotTables = protection.PivotTables;
                sheetProtection.Sort = protection.Sort;
                sheetProtection.SelectLockedCells = !protection.SelectLockedCells;
                sheetProtection.SelectUnlockedCells = !protection.SelectUnlockedCells;
            }
            else
            {
                worksheetPart.Worksheet.RemoveAllChildren<SheetProtection>();
                cm.SetElement(XLWSContentManager.XLWSContents.SheetProtection, null);
            }
            #endregion

            #region AutoFilter
            if (xlWorksheet.AutoFilterRange != null)
            {
                if (worksheetPart.Worksheet.Elements<AutoFilter>().Count() == 0)
                {
                    OpenXmlElement previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.AutoFilter);
                    worksheetPart.Worksheet.InsertAfter(new AutoFilter(), previousElement);
                }

                var autoFilter = worksheetPart.Worksheet.Elements<AutoFilter>().First();
                cm.SetElement(XLWSContentManager.XLWSContents.AutoFilter, autoFilter);

                autoFilter.Reference = xlWorksheet.AutoFilterRange.RangeAddress.ToString();
            }
            else
            {
                worksheetPart.Worksheet.RemoveAllChildren<AutoFilter>();
                cm.SetElement(XLWSContentManager.XLWSContents.AutoFilter, null);
            }
            #endregion

            #region MergeCells
            MergeCells mergeCells = null;
            if (xlWorksheet.Internals.MergedRanges.Count() > 0)
            {
                if (worksheetPart.Worksheet.Elements<MergeCells>().Count() == 0)
                {
                    OpenXmlElement previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.MergeCells);
                    worksheetPart.Worksheet.InsertAfter(new MergeCells(), previousElement);
                }

                mergeCells = worksheetPart.Worksheet.Elements<MergeCells>().First();
                cm.SetElement(XLWSContentManager.XLWSContents.MergeCells, mergeCells);
                mergeCells.RemoveAllChildren<MergeCell>();

                foreach (var merged in xlWorksheet.Internals.MergedRanges.Select(m => m.RangeAddress.FirstAddress.ToString() + ":" + m.RangeAddress.LastAddress.ToString()))
                {
                    MergeCell mergeCell = new MergeCell() { Reference = merged };
                    mergeCells.Append(mergeCell);
                }

                mergeCells.Count = (UInt32)mergeCells.Count();
            }
            else
            {
                worksheetPart.Worksheet.RemoveAllChildren<MergeCells>();
                cm.SetElement(XLWSContentManager.XLWSContents.MergeCells, null);
            }
            #endregion

            #region DataValidations
            DataValidations dataValidations = null;
            
            if (xlWorksheet.DataValidations.Count() == 0)
            {
                worksheetPart.Worksheet.RemoveAllChildren<DataValidations>();
                cm.SetElement(XLWSContentManager.XLWSContents.DataValidations, null);
            }
            else
            {
                worksheetPart.Worksheet.Elements<DataValidations>().FirstOrDefault();
                if (worksheetPart.Worksheet.Elements<DataValidations>().Count() == 0)
                {
                    OpenXmlElement previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.DataValidations);
                    worksheetPart.Worksheet.InsertAfter(new DataValidations(), previousElement);
                }

                dataValidations = worksheetPart.Worksheet.Elements<DataValidations>().First();
                cm.SetElement(XLWSContentManager.XLWSContents.DataValidations, dataValidations);
                dataValidations.RemoveAllChildren<DataValidation>();
                foreach (var dv in xlWorksheet.DataValidations)
                {
                    String sequence = String.Empty;
                    foreach (var r in dv.Ranges)
                    {
                        sequence += r.RangeAddress.ToString() + " ";
                    }

                    if (sequence.Length > 0)
                        sequence = sequence.Substring(0, sequence.Length - 1);

                    DataValidation dataValidation = new DataValidation()
                    {
                        AllowBlank = dv.IgnoreBlanks,
                        Formula1 = new Formula1(dv.MinValue),
                        Formula2 = new Formula2(dv.MaxValue),
                        Type = dataValidationValues.Single(p => p.Key == dv.AllowedValues).Value,
                        ShowErrorMessage = dv.ShowErrorMessage,
                        Prompt = dv.InputMessage,
                        PromptTitle = dv.InputTitle,
                        ErrorTitle = dv.ErrorTitle,
                        Error = dv.ErrorMessage,
                        ShowDropDown = !dv.InCellDropdown,
                        ShowInputMessage = dv.ShowInputMessage,
                        ErrorStyle = dataValidationErrorStyleValues.Single(p => p.Key == dv.ErrorStyle).Value,
                        Operator = dataValidationOperatorValues.Single(p => p.Key == dv.Operator).Value,
                        SequenceOfReferences = new ListValue<StringValue>() { InnerText = sequence }
                    };
                    
                    dataValidations.Append(dataValidation);
                }
                dataValidations.Count = (UInt32)xlWorksheet.DataValidations.Count();
            }


            #endregion

            #region Hyperlinks
            Hyperlinks hyperlinks = null;
            var relToRemove = worksheetPart.HyperlinkRelationships.ToList();
            relToRemove.ForEach(h => worksheetPart.DeleteReferenceRelationship(h));
            if (xlWorksheet.Hyperlinks.Count() == 0)
            {
                worksheetPart.Worksheet.RemoveAllChildren<Hyperlinks>();
                cm.SetElement(XLWSContentManager.XLWSContents.Hyperlinks, null);
            }
            else
            {
                worksheetPart.Worksheet.Elements<Hyperlinks>().FirstOrDefault();
                if (worksheetPart.Worksheet.Elements<Hyperlinks>().Count() == 0)
                {
                    OpenXmlElement previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.Hyperlinks);
                    worksheetPart.Worksheet.InsertAfter(new Hyperlinks(), previousElement);
                }

                hyperlinks = worksheetPart.Worksheet.Elements<Hyperlinks>().First();
                cm.SetElement(XLWSContentManager.XLWSContents.Hyperlinks, hyperlinks);
                hyperlinks.RemoveAllChildren<Hyperlink>();
                foreach (var hl in xlWorksheet.Hyperlinks)
                {
                    Hyperlink hyperlink;
                    if (hl.IsExternal)
                    {
                        String rId = relId.GetNext(RelType.Workbook);
                        hyperlink = new Hyperlink() { Reference = hl.Cell.Address.ToString(), Id = rId };
                        worksheetPart.AddHyperlinkRelationship(hl.ExternalAddress, true, rId);
                    }
                    else
                    {
                        hyperlink = new Hyperlink()
                        {
                            Reference = hl.Cell.Address.ToString(),
                            Location = hl.InternalAddress,
                            Display = hl.Cell.GetFormattedString()
                        };
                    }
                    if (!StringExtensions.IsNullOrWhiteSpace(hl.Tooltip))
                        hyperlink.Tooltip = hl.Tooltip;
                    hyperlinks.Append(hyperlink);
                }
            }
            #endregion

            #region PrintOptions
            PrintOptions printOptions = null;
            if (worksheetPart.Worksheet.Elements<PrintOptions>().Count() == 0)
            {
                OpenXmlElement previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.PrintOptions);
                worksheetPart.Worksheet.InsertAfter(new PrintOptions(), previousElement);
            }

            printOptions = worksheetPart.Worksheet.Elements<PrintOptions>().First();
            cm.SetElement(XLWSContentManager.XLWSContents.PrintOptions, printOptions);

            printOptions.HorizontalCentered = xlWorksheet.PageSetup.CenterHorizontally;
            printOptions.VerticalCentered = xlWorksheet.PageSetup.CenterVertically;
            printOptions.Headings = xlWorksheet.PageSetup.ShowRowAndColumnHeadings;
            printOptions.GridLines = xlWorksheet.PageSetup.ShowGridlines;
            #endregion

            #region PageMargins
            if (worksheetPart.Worksheet.Elements<PageMargins>().Count() == 0)
            {
                OpenXmlElement previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.PageMargins);
                worksheetPart.Worksheet.InsertAfter(new PageMargins(), previousElement);
            }

            PageMargins pageMargins = worksheetPart.Worksheet.Elements<PageMargins>().First();
            cm.SetElement(XLWSContentManager.XLWSContents.PageMargins, pageMargins);
            pageMargins.Left = xlWorksheet.PageSetup.Margins.Left;
            pageMargins.Right = xlWorksheet.PageSetup.Margins.Right;
            pageMargins.Top = xlWorksheet.PageSetup.Margins.Top;
            pageMargins.Bottom = xlWorksheet.PageSetup.Margins.Bottom;
            pageMargins.Header = xlWorksheet.PageSetup.Margins.Header;
            pageMargins.Footer = xlWorksheet.PageSetup.Margins.Footer;
            #endregion

            #region PageSetup
            if (worksheetPart.Worksheet.Elements<PageSetup>().Count() == 0)
            {
                var previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.PageSetup);
                worksheetPart.Worksheet.InsertAfter(new PageSetup(), previousElement);
            }

            PageSetup pageSetup = worksheetPart.Worksheet.Elements<PageSetup>().First();
            cm.SetElement(XLWSContentManager.XLWSContents.PageSetup, pageSetup);

            pageSetup.Orientation = pageOrientationValues.Single(p => p.Key == xlWorksheet.PageSetup.PageOrientation).Value;
            pageSetup.PaperSize = (UInt32)xlWorksheet.PageSetup.PaperSize;
            pageSetup.BlackAndWhite = xlWorksheet.PageSetup.BlackAndWhite;
            pageSetup.Draft = xlWorksheet.PageSetup.DraftQuality;
            pageSetup.PageOrder = pageOrderValues.Single(p => p.Key == xlWorksheet.PageSetup.PageOrder).Value;
            pageSetup.CellComments = showCommentsValues.Single(s => s.Key == xlWorksheet.PageSetup.ShowComments).Value;
            pageSetup.Errors = printErrorValues.Single(p => p.Key == xlWorksheet.PageSetup.PrintErrorValue).Value;

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

                if (xlWorksheet.PageSetup.PagesWide > 0)
                    pageSetup.FitToWidth = (UInt32)xlWorksheet.PageSetup.PagesWide;
                else
                    pageSetup.FitToWidth = 0;

                if (xlWorksheet.PageSetup.PagesTall > 0)
                    pageSetup.FitToHeight = (UInt32)xlWorksheet.PageSetup.PagesTall;
                else
                    pageSetup.FitToHeight = 0;
            }
            #endregion

            #region HeaderFooter
            if (worksheetPart.Worksheet.Elements<HeaderFooter>().Count() == 0)
            {
                var previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.HeaderFooter);
                worksheetPart.Worksheet.InsertAfter(new HeaderFooter(), previousElement);
            }

            HeaderFooter headerFooter = worksheetPart.Worksheet.Elements<HeaderFooter>().First();
            cm.SetElement(XLWSContentManager.XLWSContents.HeaderFooter, headerFooter);
            headerFooter.RemoveAllChildren();

            headerFooter.ScaleWithDoc = xlWorksheet.PageSetup.ScaleHFWithDocument;
            headerFooter.AlignWithMargins = xlWorksheet.PageSetup.AlignHFWithMargins;
            headerFooter.DifferentFirst = true;
            headerFooter.DifferentOddEven = true;

            OddHeader oddHeader = new OddHeader(xlWorksheet.PageSetup.Header.GetText(XLHFOccurrence.OddPages));
            headerFooter.Append(oddHeader);
            OddFooter oddFooter = new OddFooter(xlWorksheet.PageSetup.Footer.GetText(XLHFOccurrence.OddPages));
            headerFooter.Append(oddFooter);

            EvenHeader evenHeader = new EvenHeader(xlWorksheet.PageSetup.Header.GetText(XLHFOccurrence.EvenPages));
            headerFooter.Append(evenHeader);
            EvenFooter evenFooter = new EvenFooter(xlWorksheet.PageSetup.Footer.GetText(XLHFOccurrence.EvenPages));
            headerFooter.Append(evenFooter);

            FirstHeader firstHeader = new FirstHeader(xlWorksheet.PageSetup.Header.GetText(XLHFOccurrence.FirstPage));
            headerFooter.Append(firstHeader);
            FirstFooter firstFooter = new FirstFooter(xlWorksheet.PageSetup.Footer.GetText(XLHFOccurrence.FirstPage));
            headerFooter.Append(firstFooter);

            //if (!headerFooter.Any(hf => hf.InnerText.Length > 0))
            //    worksheetPart.Worksheet.RemoveAllChildren<HeaderFooter>();
            #endregion

            #region RowBreaks
            if (worksheetPart.Worksheet.Elements<RowBreaks>().Count() == 0)
            {
                OpenXmlElement previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.RowBreaks);
                worksheetPart.Worksheet.InsertAfter(new RowBreaks(), previousElement);
            }

            RowBreaks rowBreaks = worksheetPart.Worksheet.Elements<RowBreaks>().First();
            
            var rowBreakCount = xlWorksheet.PageSetup.RowBreaks.Count;
            if (rowBreakCount > 0)
            {
                rowBreaks.Count = (UInt32)rowBreakCount;
                rowBreaks.ManualBreakCount = (UInt32)rowBreakCount;
                foreach (var rb in xlWorksheet.PageSetup.RowBreaks)
                {
                    Break break1 = new Break() { Id = (UInt32)rb, Max = (UInt32)xlWorksheet.RangeAddress.LastAddress.RowNumber, ManualPageBreak = true };
                    rowBreaks.Append(break1);
                }
                cm.SetElement(XLWSContentManager.XLWSContents.RowBreaks, rowBreaks);
            }
            else
            {
                worksheetPart.Worksheet.RemoveAllChildren<RowBreaks>();
                cm.SetElement(XLWSContentManager.XLWSContents.RowBreaks, null);
            }
            #endregion

            #region ColumnBreaks

            if (worksheetPart.Worksheet.Elements<ColumnBreaks>().Count() == 0)
            {
                OpenXmlElement previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.ColumnBreaks);
                worksheetPart.Worksheet.InsertAfter(new ColumnBreaks(), previousElement);
            }

            ColumnBreaks columnBreaks = worksheetPart.Worksheet.Elements<ColumnBreaks>().First();

            var columnBreakCount = xlWorksheet.PageSetup.ColumnBreaks.Count;
            if (columnBreakCount > 0)
            {
                columnBreaks.Count = (UInt32)columnBreakCount;
                columnBreaks.ManualBreakCount = (UInt32)columnBreakCount;
                foreach (var cb in xlWorksheet.PageSetup.ColumnBreaks)
                {
                    Break break1 = new Break() { Id = (UInt32)cb, Max = (UInt32)xlWorksheet.RangeAddress.LastAddress.ColumnNumber, ManualPageBreak = true };
                    columnBreaks.Append(break1);
                }
                cm.SetElement(XLWSContentManager.XLWSContents.ColumnBreaks, columnBreaks);
            }
            else
            {
                worksheetPart.Worksheet.RemoveAllChildren<ColumnBreaks>();
                cm.SetElement(XLWSContentManager.XLWSContents.ColumnBreaks, null);
            }
            #endregion

            #region Drawings
            //worksheetPart.Worksheet.RemoveAllChildren<Drawing>();
            //{
            //    OpenXmlElement previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.Drawing);
            //    worksheetPart.Worksheet.InsertAfter(new Drawing() { Id = String.Format("rId{0}", 1) }, previousElement);
            //}

            //Drawing drawing = worksheetPart.Worksheet.Elements<Drawing>().First();
            //cm.SetElement(XLWSContentManager.XLWSContents.Drawing, drawing);
            #endregion

            #region Tables
            worksheetPart.Worksheet.RemoveAllChildren<TableParts>();
            {
                OpenXmlElement previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.TableParts);
                worksheetPart.Worksheet.InsertAfter(new TableParts(), previousElement);
            }

            TableParts tableParts = worksheetPart.Worksheet.Elements<TableParts>().First();
            cm.SetElement(XLWSContentManager.XLWSContents.TableParts, tableParts);

            tableParts.Count = (UInt32)xlWorksheet.Tables.Count();
            foreach (var table in xlWorksheet.Tables)
            {
                var xlTable = (XLTable)table;
                TablePart tablePart = new TablePart() { Id = xlTable.RelId };
                tableParts.Append(tablePart);
            }

            #endregion
        }

        private void CollapseColumns(Columns columns, Dictionary<uint, Column> sheetColumns)
        {
            UInt32 lastMax = 1;
            UInt32 lastMin = 1;
            Int32 count = sheetColumns.Count;
            foreach (var kp in sheetColumns.OrderBy(kp => kp.Key))
            {
                if (kp.Key < count && ColumnsAreEqual(kp.Value, sheetColumns[kp.Key + 1]))
                {
                    lastMax = kp.Key;
                }
                else
                {
                    var newColumn = (Column)kp.Value.CloneNode(true);
                    newColumn.Min = lastMin;
                    var columnsToRemove = new List<Column>();
                    foreach (var c in columns.Elements<Column>().Where(co => co.Min >= newColumn.Min && co.Max <= newColumn.Max).Select(co => co))
                    {
                        columnsToRemove.Add(c);
                    }
                    columnsToRemove.ForEach(c => columns.RemoveChild(c));

                    columns.Append(newColumn);
                    
                    lastMin = kp.Key + 1;
                }
                
            }
        }

        private Double GetColumnWidth(Double columnWidth)
        {
            if (columnWidth > 0)
                return columnWidth + 0.71;
            else
                return columnWidth;
        }

        private void UpdateColumn(Column column, Columns columns, Dictionary<UInt32, Column> sheetColumnsByMin)//, Dictionary<UInt32, Column> sheetColumnsByMax)
        {
            UInt32 co = column.Min.Value;
            Column newColumn;
            Column existingColumn; // = columns.Elements<Column>().FirstOrDefault(c => c.Min.Value == column.Min.Value);
            if (!sheetColumnsByMin.ContainsKey(co))
            {
                //if (sheetColumnsByMin.ContainsKey(co + 1) && ColumnsAreEqual(column, sheetColumnsByMin[co + 1]))
                //{
                //    var thisColumn = sheetColumnsByMin[co + 1];
                //    thisColumn.Min -= 1;
                //    sheetColumnsByMin.Remove(co + 1);
                //    sheetColumnsByMin.Add(co, thisColumn);
                //}
                //else if (sheetColumnsByMax.ContainsKey(co - 1) && ColumnsAreEqual(column, sheetColumnsByMin[co - 1]))
                //{
                //    var thisColumn = sheetColumnsByMin[co - 1];
                //    thisColumn.Max += 1;
                //    sheetColumnsByMax.Remove(co - 1);
                //    sheetColumnsByMax.Add(co, thisColumn);
                //}
                //else
                //{
                newColumn = (Column)column.CloneNode(true);
                columns.Append(newColumn);
                sheetColumnsByMin.Add(co, newColumn);
                //    sheetColumnsByMax.Add(co, newColumn);
                //}
            }
            else
            {
                existingColumn = sheetColumnsByMin[column.Min.Value];
                newColumn = (Column)existingColumn.CloneNode(true);
                //newColumn = new Column() { InnerXml = existingColumn.InnerXml };
                newColumn.Min = column.Min;
                newColumn.Max = column.Max;
                newColumn.Style = column.Style;
                newColumn.Width = column.Width;
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
                    columns.Append(newColumn);
                    sheetColumnsByMin.Add(newColumn.Min.Value, newColumn);
                }
                else
                {
                    //columns.InsertBefore(existingColumn, newColumn);
                    columns.Append(newColumn);
                    sheetColumnsByMin.Add(newColumn.Min.Value, newColumn);
                    existingColumn.Min = existingColumn.Min + 1;
                    sheetColumnsByMin.Add(existingColumn.Min.Value, existingColumn);
                }
            }

        }

        private bool ColumnsAreEqual(Column column, Column column_2)
        {
            return
                column.Style.Value == column_2.Style.Value
                && column.Width.Value == column_2.Width.Value
                && ((column.Hidden == null && column_2.Hidden == null)
                    || (column.Hidden != null && column_2.Hidden != null && column.Hidden.Value == column_2.Hidden.Value))
                && ((column.Collapsed == null && column_2.Collapsed == null)
                    || (column.Collapsed != null && column_2.Collapsed != null && column.Collapsed.Value == column_2.Collapsed.Value))
                && ((column.OutlineLevel == null && column_2.OutlineLevel == null)
                    || (column.OutlineLevel != null && column_2.OutlineLevel != null && column.OutlineLevel.Value == column_2.OutlineLevel.Value));
        }
        #endregion

        private void GenerateCalculationChainPartContent(WorkbookPart workbookPart)
        {
            var thisRelId = relId.GetNext(RelType.Workbook);
            if (workbookPart.CalculationChainPart == null)
                workbookPart.AddNewPart<CalculationChainPart>(thisRelId);

            if (workbookPart.CalculationChainPart.CalculationChain == null)
                workbookPart.CalculationChainPart.CalculationChain = new CalculationChain();

            CalculationChain calculationChain =  workbookPart.CalculationChainPart.CalculationChain;
            calculationChain.RemoveAllChildren<CalculationCell>();
            //var calculationCells = new Dictionary<String, List<CalculationCell>>();
            //foreach(var calculationCell in calculationChain.Elements<CalculationCell>().Where(cc => cc.CellReference != null))
            //{
            //    String cellReference = calculationCell.CellReference.Value;
            //    if (!calculationCells.ContainsKey(cellReference))
            //        calculationCells.Add(cellReference, new List<CalculationCell>());
            //    calculationCell.
            //    calculationCells[cellReference].Add(calculationCell);
            //}
                
            foreach (var worksheet in Worksheets.Cast<XLWorksheet>())
            {
                var cellsWithoutFormulas = new HashSet<String>();
                foreach (var c in worksheet.Internals.CellsCollection.Values)
                {
                    if (StringExtensions.IsNullOrWhiteSpace(c.FormulaA1))
                    {
                        cellsWithoutFormulas.Add(c.Address.ToStringRelative());
                    }
                    else
                    {
                        //var calculationCells = calculationChain.Elements<CalculationCell>().Where(
                        //    cc => cc.CellReference != null && cc.CellReference == c.Address.ToString()).Select(cc => cc).ToList();

                        //calculationCells.ForEach(cc => calculationChain.RemoveChild(cc));


                        if (c.FormulaA1.StartsWith("{"))
                        {
                            calculationChain.Append(new CalculationCell() { CellReference = c.Address.ToString(), SheetId = worksheet.SheetId, Array = true });
                            calculationChain.Append(new CalculationCell() { CellReference = c.Address.ToString(), InChildChain = true });
                        }
                        else
                        {
                            calculationChain.Append(new CalculationCell() { CellReference = c.Address.ToString(), SheetId = worksheet.SheetId });
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
                m.ToList().ForEach(cc=>calculationChain.RemoveChild(cc));
            }

            if (calculationChain.Count() == 0)
            {
                workbookPart.DeletePart(workbookPart.CalculationChainPart);
            }
        }

        private void GenerateThemePartContent(ThemePart themePart)
        {
            A.Theme theme1 = new A.Theme() { Name = "Office Theme" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements1 = new A.ThemeElements();

            A.ColorScheme colorScheme1 = new A.ColorScheme() { Name = "Office" };

            A.Dark1Color dark1Color1 = new A.Dark1Color();
            A.SystemColor systemColor1 = new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = Theme.Text1.Color.ToHex().Substring(2)};

            dark1Color1.Append(systemColor1);

            A.Light1Color light1Color1 = new A.Light1Color();
            A.SystemColor systemColor2 = new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = Theme.Background1.Color.ToHex().Substring(2) };

            light1Color1.Append(systemColor2);

            A.Dark2Color dark2Color1 = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = Theme.Text2.Color.ToHex().Substring(2) };

            dark2Color1.Append(rgbColorModelHex1);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = Theme.Background2.Color.ToHex().Substring(2) };

            light2Color1.Append(rgbColorModelHex2);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = Theme.Accent1.Color.ToHex().Substring(2) };

            accent1Color1.Append(rgbColorModelHex3);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = Theme.Accent2.Color.ToHex().Substring(2) };

            accent2Color1.Append(rgbColorModelHex4);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = Theme.Accent3.Color.ToHex().Substring(2) };

            accent3Color1.Append(rgbColorModelHex5);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = Theme.Accent4.Color.ToHex().Substring(2) };

            accent4Color1.Append(rgbColorModelHex6);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = Theme.Accent5.Color.ToHex().Substring(2) };

            accent5Color1.Append(rgbColorModelHex7);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = Theme.Accent6.Color.ToHex().Substring(2) };

            accent6Color1.Append(rgbColorModelHex8);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = Theme.Hyperlink.Color.ToHex().Substring(2) };

            hyperlink1.Append(rgbColorModelHex9);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = Theme.FollowedHyperlink.Color.ToHex().Substring(2) };

            followedHyperlinkColor1.Append(rgbColorModelHex10);

            colorScheme1.Append(dark1Color1);
            colorScheme1.Append(light1Color1);
            colorScheme1.Append(dark2Color1);
            colorScheme1.Append(light2Color1);
            colorScheme1.Append(accent1Color1);
            colorScheme1.Append(accent2Color1);
            colorScheme1.Append(accent3Color1);
            colorScheme1.Append(accent4Color1);
            colorScheme1.Append(accent5Color1);
            colorScheme1.Append(accent6Color1);
            colorScheme1.Append(hyperlink1);
            colorScheme1.Append(followedHyperlinkColor1);

            A.FontScheme fontScheme2 = new A.FontScheme() { Name = "Office" };

            A.MajorFont majorFont1 = new A.MajorFont();
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "Cambria" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont1 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ Ｐゴシック" };
            A.SupplementalFont supplementalFont2 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont3 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont4 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont5 = new A.SupplementalFont() { Script = "Arab", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont6 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont7 = new A.SupplementalFont() { Script = "Thai", Typeface = "Tahoma" };
            A.SupplementalFont supplementalFont8 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont9 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont10 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont11 = new A.SupplementalFont() { Script = "Khmr", Typeface = "MoolBoran" };
            A.SupplementalFont supplementalFont12 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont13 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont14 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont15 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont16 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont17 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont18 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont19 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont20 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont21 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont22 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont23 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont24 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont25 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont26 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont27 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont28 = new A.SupplementalFont() { Script = "Viet", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont29 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };

            majorFont1.Append(latinFont1);
            majorFont1.Append(eastAsianFont1);
            majorFont1.Append(complexScriptFont1);
            majorFont1.Append(supplementalFont1);
            majorFont1.Append(supplementalFont2);
            majorFont1.Append(supplementalFont3);
            majorFont1.Append(supplementalFont4);
            majorFont1.Append(supplementalFont5);
            majorFont1.Append(supplementalFont6);
            majorFont1.Append(supplementalFont7);
            majorFont1.Append(supplementalFont8);
            majorFont1.Append(supplementalFont9);
            majorFont1.Append(supplementalFont10);
            majorFont1.Append(supplementalFont11);
            majorFont1.Append(supplementalFont12);
            majorFont1.Append(supplementalFont13);
            majorFont1.Append(supplementalFont14);
            majorFont1.Append(supplementalFont15);
            majorFont1.Append(supplementalFont16);
            majorFont1.Append(supplementalFont17);
            majorFont1.Append(supplementalFont18);
            majorFont1.Append(supplementalFont19);
            majorFont1.Append(supplementalFont20);
            majorFont1.Append(supplementalFont21);
            majorFont1.Append(supplementalFont22);
            majorFont1.Append(supplementalFont23);
            majorFont1.Append(supplementalFont24);
            majorFont1.Append(supplementalFont25);
            majorFont1.Append(supplementalFont26);
            majorFont1.Append(supplementalFont27);
            majorFont1.Append(supplementalFont28);
            majorFont1.Append(supplementalFont29);

            A.MinorFont minorFont1 = new A.MinorFont();
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "Calibri" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont30 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ Ｐゴシック" };
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont32 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont33 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont34 = new A.SupplementalFont() { Script = "Arab", Typeface = "Arial" };
            A.SupplementalFont supplementalFont35 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Arial" };
            A.SupplementalFont supplementalFont36 = new A.SupplementalFont() { Script = "Thai", Typeface = "Tahoma" };
            A.SupplementalFont supplementalFont37 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont38 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont39 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont40 = new A.SupplementalFont() { Script = "Khmr", Typeface = "DaunPenh" };
            A.SupplementalFont supplementalFont41 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont42 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont43 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont44 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont45 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont46 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont47 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont48 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont49 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont50 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont51 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont52 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont53 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont54 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont55 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont56 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont57 = new A.SupplementalFont() { Script = "Viet", Typeface = "Arial" };
            A.SupplementalFont supplementalFont58 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };

            minorFont1.Append(latinFont2);
            minorFont1.Append(eastAsianFont2);
            minorFont1.Append(complexScriptFont2);
            minorFont1.Append(supplementalFont30);
            minorFont1.Append(supplementalFont31);
            minorFont1.Append(supplementalFont32);
            minorFont1.Append(supplementalFont33);
            minorFont1.Append(supplementalFont34);
            minorFont1.Append(supplementalFont35);
            minorFont1.Append(supplementalFont36);
            minorFont1.Append(supplementalFont37);
            minorFont1.Append(supplementalFont38);
            minorFont1.Append(supplementalFont39);
            minorFont1.Append(supplementalFont40);
            minorFont1.Append(supplementalFont41);
            minorFont1.Append(supplementalFont42);
            minorFont1.Append(supplementalFont43);
            minorFont1.Append(supplementalFont44);
            minorFont1.Append(supplementalFont45);
            minorFont1.Append(supplementalFont46);
            minorFont1.Append(supplementalFont47);
            minorFont1.Append(supplementalFont48);
            minorFont1.Append(supplementalFont49);
            minorFont1.Append(supplementalFont50);
            minorFont1.Append(supplementalFont51);
            minorFont1.Append(supplementalFont52);
            minorFont1.Append(supplementalFont53);
            minorFont1.Append(supplementalFont54);
            minorFont1.Append(supplementalFont55);
            minorFont1.Append(supplementalFont56);
            minorFont1.Append(supplementalFont57);
            minorFont1.Append(supplementalFont58);

            fontScheme2.Append(majorFont1);
            fontScheme2.Append(minorFont1);

            A.FormatScheme formatScheme1 = new A.FormatScheme() { Name = "Office" };

            A.FillStyleList fillStyleList1 = new A.FillStyleList();

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill1.Append(schemeColor1);

            A.GradientFill gradientFill1 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint1 = new A.Tint() { Val = 50000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation() { Val = 300000 };

            schemeColor2.Append(tint1);
            schemeColor2.Append(saturationModulation1);

            gradientStop1.Append(schemeColor2);

            A.GradientStop gradientStop2 = new A.GradientStop() { Position = 35000 };

            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint2 = new A.Tint() { Val = 37000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation() { Val = 300000 };

            schemeColor3.Append(tint2);
            schemeColor3.Append(saturationModulation2);

            gradientStop2.Append(schemeColor3);

            A.GradientStop gradientStop3 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint3 = new A.Tint() { Val = 15000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation() { Val = 350000 };

            schemeColor4.Append(tint3);
            schemeColor4.Append(saturationModulation3);

            gradientStop3.Append(schemeColor4);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill() { Angle = 16200000, Scaled = true };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            A.GradientFill gradientFill2 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade1 = new A.Shade() { Val = 51000 };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation() { Val = 130000 };

            schemeColor5.Append(shade1);
            schemeColor5.Append(saturationModulation4);

            gradientStop4.Append(schemeColor5);

            A.GradientStop gradientStop5 = new A.GradientStop() { Position = 80000 };

            A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade2 = new A.Shade() { Val = 93000 };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation() { Val = 130000 };

            schemeColor6.Append(shade2);
            schemeColor6.Append(saturationModulation5);

            gradientStop5.Append(schemeColor6);

            A.GradientStop gradientStop6 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade3 = new A.Shade() { Val = 94000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation() { Val = 135000 };

            schemeColor7.Append(shade3);
            schemeColor7.Append(saturationModulation6);

            gradientStop6.Append(schemeColor7);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill() { Angle = 16200000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill1);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline1 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill2 = new A.SolidFill();

            A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade4 = new A.Shade() { Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 105000 };

            schemeColor8.Append(shade4);
            schemeColor8.Append(saturationModulation7);

            solidFill2.Append(schemeColor8);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline1.Append(solidFill2);
            outline1.Append(presetDash1);

            A.Outline outline2 = new A.Outline() { Width = 25400, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill3.Append(schemeColor9);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline2.Append(solidFill3);
            outline2.Append(presetDash2);

            A.Outline outline3 = new A.Outline() { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill4.Append(schemeColor10);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline3.Append(solidFill4);
            outline3.Append(presetDash3);

            lineStyleList1.Append(outline1);
            lineStyleList1.Append(outline2);
            lineStyleList1.Append(outline3);

            A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

            A.EffectStyle effectStyle1 = new A.EffectStyle();

            A.EffectList effectList1 = new A.EffectList();

            A.OuterShadow outerShadow1 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha1 = new A.Alpha() { Val = 38000 };

            rgbColorModelHex11.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex11);

            effectList1.Append(outerShadow1);

            effectStyle1.Append(effectList1);

            A.EffectStyle effectStyle2 = new A.EffectStyle();

            A.EffectList effectList2 = new A.EffectList();

            A.OuterShadow outerShadow2 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex12 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha2 = new A.Alpha() { Val = 35000 };

            rgbColorModelHex12.Append(alpha2);

            outerShadow2.Append(rgbColorModelHex12);

            effectList2.Append(outerShadow2);

            effectStyle2.Append(effectList2);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList3 = new A.EffectList();

            A.OuterShadow outerShadow3 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex13 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha3 = new A.Alpha() { Val = 35000 };

            rgbColorModelHex13.Append(alpha3);

            outerShadow3.Append(rgbColorModelHex13);

            effectList3.Append(outerShadow3);

            A.Scene3DType scene3DType1 = new A.Scene3DType();

            A.Camera camera1 = new A.Camera() { Preset = A.PresetCameraValues.OrthographicFront };
            A.Rotation rotation1 = new A.Rotation() { Latitude = 0, Longitude = 0, Revolution = 0 };

            camera1.Append(rotation1);

            A.LightRig lightRig1 = new A.LightRig() { Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };
            A.Rotation rotation2 = new A.Rotation() { Latitude = 0, Longitude = 0, Revolution = 1200000 };

            lightRig1.Append(rotation2);

            scene3DType1.Append(camera1);
            scene3DType1.Append(lightRig1);

            A.Shape3DType shape3DType1 = new A.Shape3DType();
            A.BevelTop bevelTop1 = new A.BevelTop() { Width = 63500L, Height = 25400L };

            shape3DType1.Append(bevelTop1);

            effectStyle3.Append(effectList3);
            effectStyle3.Append(scene3DType1);
            effectStyle3.Append(shape3DType1);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill5.Append(schemeColor11);

            A.GradientFill gradientFill3 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint4 = new A.Tint() { Val = 40000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation() { Val = 350000 };

            schemeColor12.Append(tint4);
            schemeColor12.Append(saturationModulation8);

            gradientStop7.Append(schemeColor12);

            A.GradientStop gradientStop8 = new A.GradientStop() { Position = 40000 };

            A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint5 = new A.Tint() { Val = 45000 };
            A.Shade shade5 = new A.Shade() { Val = 99000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation() { Val = 350000 };

            schemeColor13.Append(tint5);
            schemeColor13.Append(shade5);
            schemeColor13.Append(saturationModulation9);

            gradientStop8.Append(schemeColor13);

            A.GradientStop gradientStop9 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade6 = new A.Shade() { Val = 20000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation() { Val = 255000 };

            schemeColor14.Append(shade6);
            schemeColor14.Append(saturationModulation10);

            gradientStop9.Append(schemeColor14);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);

            A.PathGradientFill pathGradientFill1 = new A.PathGradientFill() { Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle1 = new A.FillToRectangle() { Left = 50000, Top = -80000, Right = 50000, Bottom = 180000 };

            pathGradientFill1.Append(fillToRectangle1);

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(pathGradientFill1);

            A.GradientFill gradientFill4 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList4 = new A.GradientStopList();

            A.GradientStop gradientStop10 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint6 = new A.Tint() { Val = 80000 };
            A.SaturationModulation saturationModulation11 = new A.SaturationModulation() { Val = 300000 };

            schemeColor15.Append(tint6);
            schemeColor15.Append(saturationModulation11);

            gradientStop10.Append(schemeColor15);

            A.GradientStop gradientStop11 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor16 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade7 = new A.Shade() { Val = 30000 };
            A.SaturationModulation saturationModulation12 = new A.SaturationModulation() { Val = 200000 };

            schemeColor16.Append(shade7);
            schemeColor16.Append(saturationModulation12);

            gradientStop11.Append(schemeColor16);

            gradientStopList4.Append(gradientStop10);
            gradientStopList4.Append(gradientStop11);

            A.PathGradientFill pathGradientFill2 = new A.PathGradientFill() { Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle2 = new A.FillToRectangle() { Left = 50000, Top = 50000, Right = 50000, Bottom = 50000 };

            pathGradientFill2.Append(fillToRectangle2);

            gradientFill4.Append(gradientStopList4);
            gradientFill4.Append(pathGradientFill2);

            backgroundFillStyleList1.Append(solidFill5);
            backgroundFillStyleList1.Append(gradientFill3);
            backgroundFillStyleList1.Append(gradientFill4);

            formatScheme1.Append(fillStyleList1);
            formatScheme1.Append(lineStyleList1);
            formatScheme1.Append(effectStyleList1);
            formatScheme1.Append(backgroundFillStyleList1);

            themeElements1.Append(colorScheme1);
            themeElements1.Append(fontScheme2);
            themeElements1.Append(formatScheme1);
            A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();
            A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);

            themePart.Theme = theme1;
        }

        private void GenerateCustomFilePropertiesPartContent(CustomFilePropertiesPart customFilePropertiesPart1)
        {
            Op.Properties properties2 = new Op.Properties();
            properties2.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Int32 propertyId = 1;
            foreach (var p in CustomProperties)
            {
                propertyId++;
                Op.CustomDocumentProperty customDocumentProperty = new Op.CustomDocumentProperty() { FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}", PropertyId = propertyId, Name = p.Name };
                if (p.Type == XLCustomPropertyType.Text)
                {
                    Vt.VTLPWSTR vTLPWSTR1 = new Vt.VTLPWSTR();
                    vTLPWSTR1.Text = p.GetValue<String>();
                    customDocumentProperty.Append(vTLPWSTR1);
                }
                else if (p.Type == XLCustomPropertyType.Date)
                {
                    Vt.VTFileTime vTFileTime1 = new Vt.VTFileTime();
                    vTFileTime1.Text = p.GetValue<DateTime>().ToUniversalTime().ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'");
                    customDocumentProperty.Append(vTFileTime1);
                }
                else if (p.Type == XLCustomPropertyType.Number)
                {
                    Vt.VTDouble vTDouble1 = new Vt.VTDouble();
                    vTDouble1.Text = p.GetValue<Double>().ToString(CultureInfo.InvariantCulture);
                    customDocumentProperty.Append(vTDouble1);
                }
                else
                {
                    Vt.VTBool vTBool1 = new Vt.VTBool();
                    vTBool1.Text = p.GetValue<Boolean>().ToString().ToLower();
                    customDocumentProperty.Append(vTBool1);
                }
                properties2.Append(customDocumentProperty);
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

        UInt32 tableId = 0;
        HashSet<String> tableNames;
        private String GetTableName(String originalTableName)
        {
            String tableName = originalTableName.RemoveSpecialCharacters();
            String name = tableName;
            if (tableNames.Contains(name))
            {
                Int32 i = 1;
                name = tableName + i.ToStringLookup();
                while (tableNames.Contains(name))
                {
                    i++;
                    name = tableName + i.ToStringLookup();
                }
            }

            tableNames.Add(name);
            return name;
        }

        private void GenerateTableDefinitionPartContent(TableDefinitionPart tableDefinitionPart, XLTable xlTable)
        {
            tableId++;
            String reference;
            reference = xlTable.RangeAddress.FirstAddress.ToString() + ":" + xlTable.RangeAddress.LastAddress.ToString();
            String tableName = GetTableName(xlTable.Name);
            Table table = new Table() { 
                Id = tableId, 
                Name = tableName, 
                DisplayName = tableName, 
                Reference = reference };
            
            if (xlTable.ShowTotalsRow)
                table.TotalsRowCount = 1;
            else
                table.TotalsRowShown = false;

            TableColumns tableColumns1 = new TableColumns() { Count = (UInt32)xlTable.ColumnCount()  };
            UInt32 columnId = 0;
            foreach (var cell in xlTable.HeadersRow().Cells())
            {
                columnId++;
                String fieldName = cell.GetString();
                var xlField = xlTable.Field(fieldName);
                TableColumn tableColumn1 = new TableColumn() {
                    Id = (UInt32)columnId, 
                    Name = fieldName
                };
                if (xlTable.ShowTotalsRow)
                {
                    if (xlField.TotalsRowFunction != XLTotalsRowFunction.None)
                    {
                        tableColumn1.TotalsRowFunction = totalsRowFunctionValues.Single(p => p.Key == xlField.TotalsRowFunction).Value;

                        if (xlField.TotalsRowFunction == XLTotalsRowFunction.Custom)
                            tableColumn1.TotalsRowFormula = new TotalsRowFormula(xlField.TotalsRowFormulaA1);
                    }

                    if (!StringExtensions.IsNullOrWhiteSpace(xlField.TotalsRowLabel))
                        tableColumn1.TotalsRowLabel = xlField.TotalsRowLabel;
                }
                tableColumns1.Append(tableColumn1);
            }

            TableStyleInfo tableStyleInfo1 = new TableStyleInfo() {
                Name = Enum.GetName(typeof(XLTableTheme), xlTable.Theme), 
                ShowFirstColumn = xlTable.EmphasizeFirstColumn, 
                ShowLastColumn = xlTable.EmphasizeLastColumn, 
                ShowRowStripes = xlTable.ShowRowStripes, 
                ShowColumnStripes = xlTable.ShowColumnStripes
            };

            if (xlTable.ShowAutoFilter)
            {
                AutoFilter autoFilter1 = new AutoFilter();

                if (xlTable.ShowTotalsRow)
                    autoFilter1.Reference = xlTable.RangeAddress.FirstAddress.ToString() + ":" +
                        XLAddress.GetColumnLetterFromNumber(xlTable.RangeAddress.LastAddress.ColumnNumber) + (xlTable.RangeAddress.LastAddress.RowNumber - 1).ToStringLookup();
                else
                    autoFilter1.Reference = reference;

                table.Append(autoFilter1);
            }

            table.Append(tableColumns1);
            table.Append(tableStyleInfo1);

            tableDefinitionPart.Table = table;
        }

        //private void GenerateDrawingsPartContent(DrawingsPart drawingsPart, XLWorksheet worksheet)
        //{
        //    if (drawingsPart.WorksheetDrawing == null)
        //        drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing();

        //    var worksheetDrawing = drawingsPart.WorksheetDrawing;

        //    if (!worksheetDrawing.NamespaceDeclarations.Contains(new KeyValuePair<string, string>("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")))
        //        worksheetDrawing.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
        //    if (!worksheetDrawing.NamespaceDeclarations.Contains(new KeyValuePair<string, string>("a", "http://schemas.openxmlformats.org/drawingml/2006/main")))
        //        worksheetDrawing.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");


        //    foreach (var chart in worksheet.Charts.OrderBy(c => c.ZOrder).Select(c => c))
        //    {
        //        Xdr.TwoCellAnchor twoCellAnchor = new Xdr.TwoCellAnchor();
        //        worksheetDrawing.Append(twoCellAnchor);
        //        if (chart.Anchor == XLDrawingAnchor.MoveAndSizeWithCells)
        //            twoCellAnchor.EditAs = Xdr.EditAsValues.TwoCell;
        //        else if (chart.Anchor == XLDrawingAnchor.MoveWithCells)
        //            twoCellAnchor.EditAs = Xdr.EditAsValues.OneCell;
        //        else
        //            twoCellAnchor.EditAs = Xdr.EditAsValues.Absolute;

        //        if (twoCellAnchor.FromMarker == null)
        //            twoCellAnchor.FromMarker = new Xdr.FromMarker();
        //        twoCellAnchor.FromMarker.RowId = new Xdr.RowId((chart.FirstRow - 1).ToString());
        //        twoCellAnchor.FromMarker.RowOffset = new Xdr.RowOffset(chart.FirstRowOffset.ToString());
        //        twoCellAnchor.FromMarker.ColumnId = new Xdr.ColumnId((chart.FirstColumn - 1).ToString());
        //        twoCellAnchor.FromMarker.ColumnOffset = new Xdr.ColumnOffset(chart.FirstColumnOffset.ToString());

        //        if (twoCellAnchor.ToMarker == null)
        //            twoCellAnchor.ToMarker = new Xdr.ToMarker();
        //        twoCellAnchor.ToMarker.RowId = new Xdr.RowId((chart.LastRow - 1).ToString());
        //        twoCellAnchor.ToMarker.RowOffset = new Xdr.RowOffset(chart.LastRowOffset.ToString());
        //        twoCellAnchor.ToMarker.ColumnId = new Xdr.ColumnId((chart.LastColumn - 1).ToString());
        //        twoCellAnchor.ToMarker.ColumnOffset = new Xdr.ColumnOffset(chart.LastColumnOffset.ToString());



        //        Xdr.GraphicFrame graphicFrame = new Xdr.GraphicFrame();
        //        twoCellAnchor.Append(graphicFrame);

        //        if (graphicFrame.NonVisualGraphicFrameProperties == null)
        //            graphicFrame.NonVisualGraphicFrameProperties = new Xdr.NonVisualGraphicFrameProperties();

        //        if (graphicFrame.NonVisualGraphicFrameProperties.NonVisualDrawingProperties == null)
        //            graphicFrame.NonVisualGraphicFrameProperties.NonVisualDrawingProperties = new Xdr.NonVisualDrawingProperties() { Id = (UInt32)chart.Id, Name = chart.Name, Description = chart.Description, Hidden = chart.Hidden };
        //        if (graphicFrame.NonVisualGraphicFrameProperties.NonVisualGraphicFrameDrawingProperties == null)
        //            graphicFrame.NonVisualGraphicFrameProperties.NonVisualGraphicFrameDrawingProperties = new Xdr.NonVisualGraphicFrameDrawingProperties();

        //        if (graphicFrame.Transform == null)
        //            graphicFrame.Transform = new Xdr.Transform();

        //        if (chart.HorizontalFlip)
        //            graphicFrame.Transform.HorizontalFlip = true;
        //        else
        //            graphicFrame.Transform.HorizontalFlip = null;

        //        if (chart.VerticalFlip)
        //            graphicFrame.Transform.VerticalFlip = true;
        //        else
        //            graphicFrame.Transform.VerticalFlip = null;

        //        if (chart.Rotation != 0)
        //            graphicFrame.Transform.Rotation = chart.Rotation;
        //        else
        //            graphicFrame.Transform.Rotation = null;

        //        if (graphicFrame.Transform.Offset == null)
        //            graphicFrame.Transform.Offset = new A.Offset();

        //        graphicFrame.Transform.Offset.X = chart.OffsetX;
        //        graphicFrame.Transform.Offset.Y = chart.OffsetY;

        //        if (graphicFrame.Transform.Extents == null)
        //            graphicFrame.Transform.Extents = new A.Extents();

        //        graphicFrame.Transform.Extents.Cx = chart.ExtentLength;
        //        graphicFrame.Transform.Extents.Cy = chart.ExtentWidth;

        //        if (graphicFrame.Graphic == null)
        //            graphicFrame.Graphic = new A.Graphic();

        //        if (graphicFrame.Graphic.GraphicData == null)
        //            graphicFrame.Graphic.GraphicData = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" };

        //        if (!graphicFrame.Graphic.GraphicData.Elements<C.ChartReference>().Any())
        //        {
        //            C.ChartReference chartReference = new C.ChartReference() { Id = "rId" + chart.Id.ToStringLookup() };
        //            chartReference.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
        //            chartReference.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        //            graphicFrame.Graphic.GraphicData.Append(chartReference);
        //        }

        //        if (!twoCellAnchor.Elements<Xdr.ClientData>().Any())
        //            twoCellAnchor.Append(new Xdr.ClientData());
        //    }
        //}

        //private void GenerateChartPartContent(ChartPart chartPart, XLChart xlChart)
        //{
        //    if (chartPart.ChartSpace == null)
        //        chartPart.ChartSpace = new C.ChartSpace();

        //    C.ChartSpace chartSpace = chartPart.ChartSpace;

        //    if (!chartSpace.NamespaceDeclarations.Contains(new KeyValuePair<string, string>("c", "http://schemas.openxmlformats.org/drawingml/2006/chart")))
        //        chartSpace.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
        //    if (!chartSpace.NamespaceDeclarations.Contains(new KeyValuePair<string, string>("a", "http://schemas.openxmlformats.org/drawingml/2006/main")))
        //        chartSpace.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        //    if (!chartSpace.NamespaceDeclarations.Contains(new KeyValuePair<string, string>("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")))
        //        chartSpace.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        //    if (chartSpace.EditingLanguage == null)
        //        chartSpace.EditingLanguage = new C.EditingLanguage() { Val = CultureInfo.CurrentCulture.Name };
        //    else
        //        chartSpace.EditingLanguage.Val = CultureInfo.CurrentCulture.Name;

        //    C.Chart chart = new C.Chart();
        //    chartSpace.Append(chart);

        //    if (chart.Title == null)
        //        chart.Title = new C.Title();

        //    if (chart.Title.Layout == null)
        //        chart.Title.Layout = new C.Layout();

        //    if (chart.View3D == null)
        //        chart.View3D = new C.View3D();

        //    if (chart.View3D.RightAngleAxes == null)
        //        chart.View3D.RightAngleAxes = new C.RightAngleAxes();

        //    chart.View3D.RightAngleAxes.Val = xlChart.RightAngleAxes;

        //    if (chart.PlotArea == null)
        //        chart.PlotArea = new C.PlotArea();

        //    if (chart.PlotArea.Layout == null)
        //        chart.PlotArea.Layout = new C.Layout();

        //    OpenXmlElement chartElement = GetChartElement(xlChart);

        //    chart.PlotArea.Append(chartElement);



        //    C.CategoryAxis categoryAxis1 = new C.CategoryAxis();
        //    C.AxisId axisId4 = new C.AxisId() { Val = (UInt32Value)71429120U };

        //    C.Scaling scaling1 = new C.Scaling();
        //    C.Orientation orientation1 = new C.Orientation() { Val = C.OrientationValues.MinMax };

        //    scaling1.Append(orientation1);
        //    C.AxisPosition axisPosition1 = new C.AxisPosition() { Val = C.AxisPositionValues.Bottom };
        //    C.TickLabelPosition tickLabelPosition1 = new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo };
        //    C.CrossingAxis crossingAxis1 = new C.CrossingAxis() { Val = (UInt32Value)71432064U };
        //    C.Crosses crosses1 = new C.Crosses() { Val = C.CrossesValues.AutoZero };
        //    C.AutoLabeled autoLabeled1 = new C.AutoLabeled() { Val = true };
        //    C.LabelAlignment labelAlignment1 = new C.LabelAlignment() { Val = C.LabelAlignmentValues.Center };
        //    C.LabelOffset labelOffset1 = new C.LabelOffset() { Val = (UInt16Value)100U };

        //    categoryAxis1.Append(axisId4);
        //    categoryAxis1.Append(scaling1);
        //    categoryAxis1.Append(axisPosition1);
        //    categoryAxis1.Append(tickLabelPosition1);
        //    categoryAxis1.Append(crossingAxis1);
        //    categoryAxis1.Append(crosses1);
        //    categoryAxis1.Append(autoLabeled1);
        //    categoryAxis1.Append(labelAlignment1);
        //    categoryAxis1.Append(labelOffset1);

        //    C.ValueAxis valueAxis1 = new C.ValueAxis();
        //    C.AxisId axisId5 = new C.AxisId() { Val = (UInt32Value)71432064U };

        //    C.Scaling scaling2 = new C.Scaling();
        //    C.Orientation orientation2 = new C.Orientation() { Val = C.OrientationValues.MinMax };

        //    scaling2.Append(orientation2);
        //    C.AxisPosition axisPosition2 = new C.AxisPosition() { Val = C.AxisPositionValues.Left };
        //    C.MajorGridlines majorGridlines1 = new C.MajorGridlines();
        //    C.NumberingFormat numberingFormat1 = new C.NumberingFormat() { FormatCode = "General", SourceLinked = true };
        //    C.TickLabelPosition tickLabelPosition2 = new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo };
        //    C.CrossingAxis crossingAxis2 = new C.CrossingAxis() { Val = (UInt32Value)71429120U };
        //    C.Crosses crosses2 = new C.Crosses() { Val = C.CrossesValues.AutoZero };
        //    C.CrossBetween crossBetween1 = new C.CrossBetween() { Val = C.CrossBetweenValues.Between };

        //    valueAxis1.Append(axisId5);
        //    valueAxis1.Append(scaling2);
        //    valueAxis1.Append(axisPosition2);
        //    valueAxis1.Append(majorGridlines1);
        //    valueAxis1.Append(numberingFormat1);
        //    valueAxis1.Append(tickLabelPosition2);
        //    valueAxis1.Append(crossingAxis2);
        //    valueAxis1.Append(crosses2);
        //    valueAxis1.Append(crossBetween1);


        //    plotArea.Append(bar3DChart1);
        //    plotArea.Append(categoryAxis1);
        //    plotArea.Append(valueAxis1);

        //    C.Legend legend1 = new C.Legend();
        //    C.LegendPosition legendPosition1 = new C.LegendPosition() { Val = C.LegendPositionValues.Right };
        //    C.Layout layout3 = new C.Layout();

        //    legend1.Append(legendPosition1);
        //    legend1.Append(layout3);
        //    C.PlotVisibleOnly plotVisibleOnly1 = new C.PlotVisibleOnly() { Val = true };



        //    chart.Append(legend1);
        //    chart.Append(plotVisibleOnly1);

        //    C.PrintSettings printSettings1 = new C.PrintSettings();
        //    C.HeaderFooter headerFooter1 = new C.HeaderFooter();
        //    C.PageMargins pageMargins4 = new C.PageMargins() { Left = 0.70000000000000018D, Right = 0.70000000000000018D, Top = 0.75000000000000022D, Bottom = 0.75000000000000022D, Header = 0.3000000000000001D, Footer = 0.3000000000000001D };
        //    C.PageSetup pageSetup1 = new C.PageSetup();

        //    printSettings1.Append(headerFooter1);
        //    printSettings1.Append(pageMargins4);
        //    printSettings1.Append(pageSetup1);



        //    chartSpace.Append(printSettings1);

        //}

        //private OpenXmlElement GetChartElement(XLChart xlChart)
        //{
        //    if (xlChart.ChartTypeCategory == XLChartTypeCategory.Bar3D)
        //        return GetBar3DChart(xlChart);
        //    else
        //        return null;
        //}

        //private OpenXmlElement GetBar3DChart(XLChart xlChart)
        //{

        //    C.Bar3DChart bar3DChart = new C.Bar3DChart();
        //    bar3DChart.BarDirection = new C.BarDirection() { Val = GetBarDirection(xlChart) };
        //    bar3DChart.BarGrouping = new C.BarGrouping() { Val = GetBarGrouping(xlChart) };

        //    C.BarChartSeries barChartSeries = new C.BarChartSeries();
        //    barChartSeries.Index = new C.Index() { Val = (UInt32Value)0U };
        //    barChartSeries.Order = new C.Order() { Val = (UInt32Value)0U };

        //    C.SeriesText seriesText1 = new C.SeriesText();

        //    C.StringReference stringReference1 = new C.StringReference();
        //    C.Formula formula1 = new C.Formula();
        //    formula1.Text = "Sheet1!$B$1";
            
        //    stringReference1.Append(formula1);
            
        //    seriesText1.Append(stringReference1);

        //    C.CategoryAxisData categoryAxisData1 = new C.CategoryAxisData();

        //    C.StringReference stringReference2 = new C.StringReference();
        //    C.Formula formula2 = new C.Formula();
        //    formula2.Text = "Sheet1!$A$2:$A$3";

        //    C.StringCache stringCache2 = new C.StringCache();
        //    C.PointCount pointCount2 = new C.PointCount() { Val = (UInt32Value)2U };

        //    C.StringPoint stringPoint2 = new C.StringPoint() { Index = (UInt32Value)0U };
        //    C.NumericValue numericValue2 = new C.NumericValue();
        //    numericValue2.Text = "A";

        //    stringPoint2.Append(numericValue2);

        //    C.StringPoint stringPoint3 = new C.StringPoint() { Index = (UInt32Value)1U };
        //    C.NumericValue numericValue3 = new C.NumericValue();
        //    numericValue3.Text = "B";

        //    stringPoint3.Append(numericValue3);

        //    stringCache2.Append(pointCount2);
        //    stringCache2.Append(stringPoint2);
        //    stringCache2.Append(stringPoint3);

        //    stringReference2.Append(formula2);
        //    stringReference2.Append(stringCache2);

        //    categoryAxisData1.Append(stringReference2);

        //    C.Values values1 = new C.Values();

        //    C.NumberReference numberReference1 = new C.NumberReference();
        //    C.Formula formula3 = new C.Formula();
        //    formula3.Text = "Sheet1!$B$2:$B$3";

        //    C.NumberingCache numberingCache1 = new C.NumberingCache();
        //    C.FormatCode formatCode1 = new C.FormatCode();
        //    formatCode1.Text = "General";
        //    C.PointCount pointCount3 = new C.PointCount() { Val = (UInt32Value)2U };

        //    C.NumericPoint numericPoint1 = new C.NumericPoint() { Index = (UInt32Value)0U };
        //    C.NumericValue numericValue4 = new C.NumericValue();
        //    numericValue4.Text = "5";

        //    numericPoint1.Append(numericValue4);

        //    C.NumericPoint numericPoint2 = new C.NumericPoint() { Index = (UInt32Value)1U };
        //    C.NumericValue numericValue5 = new C.NumericValue();
        //    numericValue5.Text = "10";

        //    numericPoint2.Append(numericValue5);

        //    numberingCache1.Append(formatCode1);
        //    numberingCache1.Append(pointCount3);
        //    numberingCache1.Append(numericPoint1);
        //    numberingCache1.Append(numericPoint2);

        //    numberReference1.Append(formula3);
        //    numberReference1.Append(numberingCache1);

        //    values1.Append(numberReference1);

        //    barChartSeries.Append(index1);
        //    barChartSeries.Append(order1);
        //    barChartSeries.Append(seriesText1);
        //    barChartSeries.Append(categoryAxisData1);
        //    barChartSeries.Append(values1);
        //    C.Shape shape1 = new C.Shape() { Val = C.ShapeValues.Box };
        //    C.AxisId axisId1 = new C.AxisId() { Val = (UInt32Value)71429120U };
        //    C.AxisId axisId2 = new C.AxisId() { Val = (UInt32Value)71432064U };
        //    C.AxisId axisId3 = new C.AxisId() { Val = (UInt32Value)0U };


        //    bar3DChart.Append(barChartSeries);
        //    bar3DChart.Append(shape1);
        //    bar3DChart.Append(axisId1);
        //    bar3DChart.Append(axisId2);
        //    bar3DChart.Append(axisId3);

        //    return bar3DChart;
        //}

        //private C.BarGroupingValues GetBarGrouping(XLChart xlChart)
        //{
        //    if (xlChart.BarGrouping == XLBarGrouping.Clustered)
        //        return C.BarGroupingValues.Clustered;
        //    else if (xlChart.BarGrouping == XLBarGrouping.Percent)
        //        return C.BarGroupingValues.PercentStacked;
        //    else if (xlChart.BarGrouping == XLBarGrouping.Stacked)
        //        return C.BarGroupingValues.Stacked;
        //    else
        //        return C.BarGroupingValues.Standard;
        //}

        //private C.BarDirectionValues GetBarDirection(XLChart xlChart)
        //{
        //    if (xlChart.BarOrientation == XLBarOrientation.Vertical)
        //        return C.BarDirectionValues.Column;
        //    else
        //        return C.BarDirectionValues.Bar;
        //}

        
    }
}