using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using C = DocumentFormat.OpenXml.Drawing.Charts;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;



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
        private void PopulateEnums()
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
            PoulateReferenceModeValues();
        }

        private enum RelType { General, Workbook, Worksheet }
        private class RelId
        {
            private static Dictionary<RelType, Int32> relIds = new Dictionary<RelType, Int32>();
            public static Int32 GetNext(RelType relType)
            {
                if (!relIds.ContainsKey(relType))
                    relIds.Add(relType, -1);
                var relId = relIds[relType];
                relIds[relType] = ++relId;
                return relId;
            }
        }

        private Dictionary<String, UInt32> sharedStrings = new Dictionary<string, UInt32>();
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

        private Dictionary<String, StyleInfo> sharedStyles = new Dictionary<String, StyleInfo>();

        private CellValues GetCellValue(XLCellValues xlCellValue)
        {
            switch (xlCellValue)
            {
                case XLCellValues.Boolean: return CellValues.Boolean;
                case XLCellValues.DateTime: return CellValues.Date;
                case XLCellValues.Number: return CellValues.Number;
                case XLCellValues.Text: return CellValues.SharedString;
                default: throw new NotImplementedException();
            }
        }

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

        private void PoulateReferenceModeValues()
        {
            referenceModeValues.Add(new KeyValuePair<XLReferenceStyle, ReferenceModeValues>(XLReferenceStyle.R1C1, ReferenceModeValues.R1C1));
            referenceModeValues.Add(new KeyValuePair<XLReferenceStyle, ReferenceModeValues>(XLReferenceStyle.A1, ReferenceModeValues.A1));
        }

        // Creates a SpreadsheetDocument.
        private void CreatePackage(String filePath)
        {
            using (SpreadsheetDocument package = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                RelId.GetNext(RelType.Worksheet);
                CreateParts(package);
            }
        }

        // Adds child parts and generates content of the specified part.
        private void CreateParts(SpreadsheetDocument document)
        {
            Int32 startId = Worksheets.Count();
            ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId" + (startId));
            GenerateExtendedFilePropertiesPartContent(extendedFilePropertiesPart1);

            WorkbookPart workbookPart = document.AddWorkbookPart();
            GenerateWorkbookPartContent(workbookPart);

            SharedStringTablePart sharedStringTablePart = workbookPart.AddNewPart<SharedStringTablePart>("rId" + (startId + 3));
            GenerateSharedStringTablePartContent(sharedStringTablePart);

            WorkbookStylesPart workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>("rId" + (startId + 2));
            GenerateWorkbookStylesPartContent(workbookStylesPart);

            UInt32 sheetId = 0;
            foreach (var worksheet in Worksheets)
            {
                sheetId++;
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>("rId" + sheetId.ToString());
                GenerateWorksheetPartContent(worksheetPart, (XLWorksheet)worksheet);
            }

            GenerateCalculationChainPartContent(workbookPart, "rId" + (startId + 4));

            ThemePart themePart1 = workbookPart.AddNewPart<ThemePart>("rId" + (startId + 1));
            GenerateThemePartContent(themePart1);

            SetPackageProperties(document);
        }

        private void GenerateExtendedFilePropertiesPartContent(ExtendedFilePropertiesPart extendedFilePropertiesPart)
        {
            Ap.Properties properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Excel";
            Ap.DocumentSecurity documentSecurity1 = new Ap.DocumentSecurity();
            documentSecurity1.Text = "0";
            Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";

            Ap.HeadingPairs headingPairs1 = new Ap.HeadingPairs();

            Vt.VTVector vTVector1 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Variant, Size = (UInt32Value)4U };

            Vt.Variant variant1 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR1 = new Vt.VTLPSTR();
            vTLPSTR1.Text = "Worksheets";

            variant1.Append(vTLPSTR1);

            Vt.Variant variant2 = new Vt.Variant();
            Vt.VTInt32 vTInt321 = new Vt.VTInt32();
            vTInt321.Text = Worksheets.Count().ToString();

            variant2.Append(vTInt321);

            Vt.Variant variant3 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR2 = new Vt.VTLPSTR();
            vTLPSTR2.Text = "Named Ranges";

            variant3.Append(vTLPSTR2);

            var namedCount = NamedRanges.Count() + Worksheets.Aggregate(0, (counter, ws) => counter += ws.NamedRanges.Count());
            Vt.Variant variant4 = new Vt.Variant();
            Vt.VTInt32 vTInt322 = new Vt.VTInt32();
            vTInt322.Text = (
                Worksheets.Count() * 2 // for the worksheets print area and titles
                + namedCount
                ).ToString();

            variant4.Append(vTInt322);

            vTVector1.Append(variant1);
            vTVector1.Append(variant2);
            vTVector1.Append(variant3);
            vTVector1.Append(variant4);

            headingPairs1.Append(vTVector1);

            Ap.TitlesOfParts titlesOfParts1 = new Ap.TitlesOfParts();

            UInt32 sheetCount = (UInt32)Worksheets.Count();
            Vt.VTVector vTVector2 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Lpstr, Size = (UInt32Value)(sheetCount * 3 + namedCount) };
            foreach (var worksheet in Worksheets)
            {
                Vt.VTLPSTR vTLPSTR3 = new Vt.VTLPSTR();
                vTLPSTR3.Text = worksheet.Name;
                vTVector2.Append(vTLPSTR3);

                Vt.VTLPSTR vTLPSTR4 = new Vt.VTLPSTR();
                vTLPSTR4.Text = worksheet.Name + "!Print_Area";
                vTVector2.Append(vTLPSTR4);

                Vt.VTLPSTR vTLPSTR5 = new Vt.VTLPSTR();
                vTLPSTR5.Text = worksheet.Name + "!Print_Titles";
                vTVector2.Append(vTLPSTR5);

                foreach (var nr in worksheet.NamedRanges)
                {
                    Vt.VTLPSTR vTLPSTR6 = new Vt.VTLPSTR();
                    vTLPSTR6.Text = worksheet.Name + "!" + nr.Name;
                    vTVector2.Append(vTLPSTR6);
                }
            }

            foreach (var nr in NamedRanges)
            {
                Vt.VTLPSTR vTLPSTR7 = new Vt.VTLPSTR();
                vTLPSTR7.Text = nr.Name;
                vTVector2.Append(vTLPSTR7);
            }

            titlesOfParts1.Append(vTVector2);
            Ap.Manager manager1 = new Ap.Manager();
            manager1.Text = Properties.Manager;
            Ap.Company company1 = new Ap.Company();
            company1.Text = Properties.Company;

            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
            sharedDocument1.Text = "false";
            Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "12.0000";

            properties1.Append(application1);
            properties1.Append(documentSecurity1);
            properties1.Append(scaleCrop1);
            properties1.Append(headingPairs1);
            properties1.Append(titlesOfParts1);
            properties1.Append(manager1);
            properties1.Append(company1);
            properties1.Append(linksUpToDate1);
            properties1.Append(sharedDocument1);
            properties1.Append(hyperlinksChanged1);
            properties1.Append(applicationVersion1);

            extendedFilePropertiesPart.Properties = properties1;
        }

        private void GenerateWorkbookPartContent(WorkbookPart workbookPart)
        {
            Workbook workbook1 = new Workbook();
            workbook1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            FileVersion fileVersion1 = new FileVersion() { ApplicationName = "xl", LastEdited = "4", LowestEdited = "4", BuildVersion = "4506" };
            WorkbookProperties workbookProperties1 = new WorkbookProperties() { CodeName = "ThisWorkbook", DefaultThemeVersion = (UInt32Value)124226U };

            BookViews bookViews1 = new BookViews();
            WorkbookView workbookView1 = new WorkbookView() { XWindow = 0, YWindow = 30, WindowWidth = (UInt32Value)14160U, WindowHeight = (UInt32Value)11580U };

            bookViews1.Append(workbookView1);

            UInt32 sheetId = 0;
            Sheets sheets = new Sheets();
            DefinedNames definedNames = new DefinedNames();
            foreach (var worksheet in Worksheets.Cast<XLWorksheet>())
            {
                sheetId++;
                Sheet sheet = new Sheet() { Name = worksheet.Name, SheetId = (UInt32Value)sheetId, Id = "rId" + sheetId.ToString() };
                sheets.Append(sheet);

                if (worksheet.PageSetup.PrintAreas.Count() == 0)
                {
                    var minCell = worksheet.Internals.CellsCollection.Min(c => c.Key);
                    var maxCell = worksheet.Internals.CellsCollection.Max(c => c.Key);
                    if (minCell != null && maxCell != null)
                        worksheet.PageSetup.PrintAreas.Add(minCell, maxCell);
                }
                if (worksheet.PageSetup.PrintAreas.Count() > 0)
                {
                    DefinedName definedName = new DefinedName() { Name = "_xlnm.Print_Area", LocalSheetId = (UInt32Value)sheetId - 1 };
                    var definedNameText = String.Empty;
                    foreach (var printArea in worksheet.PageSetup.PrintAreas)
                    {
                        definedNameText += "'" + worksheet.Name + "'!"
                        + printArea.RangeAddress.FirstAddress.ToString()
                        + ":" + printArea.RangeAddress.LastAddress.ToString() + ",";
                    }
                    definedName.Text = definedNameText.Substring(0, definedNameText.Length - 1);
                    definedNames.Append(definedName);
                }

                foreach (var nr in worksheet.NamedRanges)
                {
                    DefinedName definedName = new DefinedName() { 
                        Name = nr.Name, 
                        LocalSheetId = (UInt32Value)sheetId - 1,
                        Text = nr.ToString()
                    };
                    if (!String.IsNullOrWhiteSpace(nr.Comment)) definedName.Comment = nr.Comment;
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
                    DefinedName definedName = new DefinedName() { Name = "_xlnm.Print_Titles", LocalSheetId = (UInt32Value)sheetId - 1 };
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
                if (!String.IsNullOrWhiteSpace(nr.Comment)) definedName.Comment = nr.Comment;
                definedNames.Append(definedName);
            }

            CalculationProperties calculationProperties = new CalculationProperties() { CalculationId = (UInt32Value)125725U };
            if (CalculateMode != XLCalculateMode.Default)
                calculationProperties.CalculationMode = calculateModeValues.Single(p => p.Key == CalculateMode).Value;

            if (ReferenceStyle != XLReferenceStyle.Default)
                calculationProperties.ReferenceMode = referenceModeValues.Single(p=>p.Key==ReferenceStyle).Value;
           

            workbook1.Append(fileVersion1);
            workbook1.Append(workbookProperties1);
            workbook1.Append(bookViews1);
            workbook1.Append(sheets);
            if (definedNames.Count() > 0) workbook1.Append(definedNames);
            workbook1.Append(calculationProperties);

            workbookPart.Workbook = workbook1;
        }

        private void GenerateSharedStringTablePartContent(SharedStringTablePart sharedStringTablePart)
        {
            List<String> combined = new List<String>();
            Worksheets.Cast<XLWorksheet>().ForEach(w => combined.AddRange(w.Internals.CellsCollection.Values.Where(c => c.DataType == XLCellValues.Text && c.InnerText != null).Select(c => c.GetString()).Distinct()));
            var distinctStrings = combined.Distinct();
            UInt32 stringCount = (UInt32)distinctStrings.Count();
            SharedStringTable sharedStringTable = new SharedStringTable() { Count = (UInt32Value)stringCount, UniqueCount = (UInt32Value)stringCount };

            UInt32 stringId = 0;
            foreach (var s in distinctStrings)
            {
                sharedStrings.Add(s, stringId++);

                SharedStringItem sharedStringItem = new SharedStringItem();
                Text text = new Text();
                text.Text = s;
                sharedStringItem.Append(text);
                sharedStringTable.Append(sharedStringItem);
            }

            sharedStringTablePart.SharedStringTable = sharedStringTable;
        }

        private void GenerateWorkbookStylesPartContent(WorkbookStylesPart workbookStylesPart)
        {
            var defaultStyle = DefaultStyle;
            Dictionary<String, FontInfo> sharedFonts = new Dictionary<String, FontInfo>();
            sharedFonts.Add(defaultStyle.Font.ToString(), new FontInfo() { FontId = 0, Font = defaultStyle.Font });

            Dictionary<String, FillInfo> sharedFills = new Dictionary<String, FillInfo>();
            sharedFills.Add(defaultStyle.Fill.ToString(), new FillInfo() { FillId = 2, Fill = defaultStyle.Fill });

            Dictionary<String, BorderInfo> sharedBorders = new Dictionary<String, BorderInfo>();
            sharedBorders.Add(defaultStyle.Border.ToString(), new BorderInfo() { BorderId = 0, Border = defaultStyle.Border });

            Dictionary<String, NumberFormatInfo> sharedNumberFormats = new Dictionary<String, NumberFormatInfo>();
            sharedNumberFormats.Add(defaultStyle.NumberFormat.ToString(), new NumberFormatInfo() { NumberFormatId = 0, NumberFormat = defaultStyle.NumberFormat });

            //Dictionary<String, AlignmentInfo> sharedAlignments = new Dictionary<String, AlignmentInfo>();
            //sharedAlignments.Add(defaultStyle.Alignment.ToString(), new AlignmentInfo() { AlignmentId = 0, Alignment = defaultStyle.Alignment });

            sharedStyles.Add(defaultStyle.ToString(),
                new StyleInfo()
                {
                    StyleId = 0,
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
                if (!sharedFonts.ContainsKey(xlStyle.Font.ToString()))
                {
                    sharedFonts.Add(xlStyle.Font.ToString(), new FontInfo() { FontId = fontCount++, Font = xlStyle.Font });
                }

                if (!sharedFills.ContainsKey(xlStyle.Fill.ToString()))
                {
                    sharedFills.Add(xlStyle.Fill.ToString(), new FillInfo() { FillId = fillCount++, Fill = xlStyle.Fill });
                }

                if (!sharedBorders.ContainsKey(xlStyle.Border.ToString()))
                {
                    sharedBorders.Add(xlStyle.Border.ToString(), new BorderInfo() { BorderId = borderCount++, Border = xlStyle.Border });
                }

                if (xlStyle.NumberFormat.NumberFormatId == -1 && !sharedNumberFormats.ContainsKey(xlStyle.NumberFormat.ToString()))
                {
                    sharedNumberFormats.Add(xlStyle.NumberFormat.ToString(), new NumberFormatInfo() { NumberFormatId = numberFormatCount + 164, NumberFormat = xlStyle.NumberFormat });
                    numberFormatCount++;
                }

                if (!sharedStyles.ContainsKey(xlStyle.ToString()))
                {
                    Int32 numberFormatId;
                    if (xlStyle.NumberFormat.NumberFormatId >= 0)
                        numberFormatId = xlStyle.NumberFormat.NumberFormatId;
                    else
                        numberFormatId = sharedNumberFormats[xlStyle.NumberFormat.ToString()].NumberFormatId;

                    sharedStyles.Add(xlStyle.ToString(),
                        new StyleInfo()
                        {
                            StyleId = styleCount++,
                            Style = xlStyle,
                            FontId = sharedFonts[xlStyle.Font.ToString()].FontId,
                            FillId = sharedFills[xlStyle.Fill.ToString()].FillId,
                            BorderId = sharedBorders[xlStyle.Border.ToString()].BorderId,
                            NumberFormatId = numberFormatId
                        });
                }
            }

            Stylesheet stylesheet1 = new Stylesheet();

            NumberingFormats numberingFormats = new NumberingFormats() { Count = (UInt32Value)(UInt32)numberFormatCount };
            foreach (var numberFormatInfo in sharedNumberFormats.Values)
            {
                NumberingFormat numberingFormat = new NumberingFormat() { NumberFormatId = (UInt32Value)(UInt32)numberFormatInfo.NumberFormatId, FormatCode = numberFormatInfo.NumberFormat.Format };
                numberingFormats.Append(numberingFormat);
            }

            Fonts fonts = new Fonts() { Count = (UInt32Value)fontCount };

            foreach (var fontInfo in sharedFonts.Values)
            {
                Bold bold = fontInfo.Font.Bold ? new Bold() : null;
                Italic italic = fontInfo.Font.Italic ? new Italic() : null;
                Underline underline = fontInfo.Font.Underline != XLFontUnderlineValues.None ? new Underline() { Val = underlineValuesList.Single(u=>u.Key == fontInfo.Font.Underline).Value } : null;
                Strike strike = fontInfo.Font.Strikethrough ? new Strike() : null;
                VerticalTextAlignment verticalAlignment = new VerticalTextAlignment() { Val = fontVerticalTextAlignmentValues.Single(f=>f.Key == fontInfo.Font.VerticalAlignment).Value };
                Shadow shadow = fontInfo.Font.Shadow ? new Shadow() : null;
                Font font = new Font();
                FontSize fontSize = new FontSize() { Val = fontInfo.Font.FontSize };
                Color color = new Color() { Rgb = fontInfo.Font.FontColor.ToHex() };
                FontName fontName = new FontName() { Val = fontInfo.Font.FontName };
                FontFamilyNumbering fontFamilyNumbering = new FontFamilyNumbering() { Val = (Int32)fontInfo.Font.FontFamilyNumbering };
                //FontScheme fontScheme = new FontScheme() { Val = FontSchemeValues.Minor };

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
                //font.Append(fontScheme);

                fonts.Append(font);
            }

            Fills fills = new Fills() { Count = (UInt32Value)fillCount };

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };
            fill1.Append(patternFill1);
            fills.Append(fill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };
            fill2.Append(patternFill2);
            fills.Append(fill2);

            foreach (var fillInfo in sharedFills.Values)
            {
                Fill fill = new Fill();
                
                PatternFill patternFill = new PatternFill() { PatternType = fillPatternValues.Single(p=>p.Key == fillInfo.Fill.PatternType).Value };
                ForegroundColor foregroundColor = new ForegroundColor() { Rgb = fillInfo.Fill.PatternColor.ToHex() };
                BackgroundColor backgroundColor = new BackgroundColor() { Rgb = fillInfo.Fill.PatternBackgroundColor.ToHex() };

                patternFill.Append(foregroundColor);
                patternFill.Append(backgroundColor);

                fill.Append(patternFill);
                fills.Append(fill);
            }

            Borders borders = new Borders() { Count = (UInt32Value)borderCount };

            foreach (var borderInfo in sharedBorders.Values)
            {
                Border border = new Border() { DiagonalUp = borderInfo.Border.DiagonalUp, DiagonalDown = borderInfo.Border.DiagonalDown };

                LeftBorder leftBorder = new LeftBorder() { Style = borderStyleValues.Single(b=>b.Key == borderInfo.Border.LeftBorder).Value };
                Color leftBorderColor = new Color() { Rgb = borderInfo.Border.LeftBorderColor.ToHex() };
                leftBorder.Append(leftBorderColor);
                border.Append(leftBorder);

                RightBorder rightBorder = new RightBorder() { Style = borderStyleValues.Single(b => b.Key == borderInfo.Border.RightBorder).Value };
                Color rightBorderColor = new Color() { Rgb = borderInfo.Border.RightBorderColor.ToHex() };
                rightBorder.Append(rightBorderColor);
                border.Append(rightBorder);

                TopBorder topBorder = new TopBorder() { Style = borderStyleValues.Single(b => b.Key == borderInfo.Border.TopBorder).Value };
                Color topBorderColor = new Color() { Rgb = borderInfo.Border.TopBorderColor.ToHex() };
                topBorder.Append(topBorderColor);
                border.Append(topBorder);

                BottomBorder bottomBorder = new BottomBorder() { Style = borderStyleValues.Single(b => b.Key == borderInfo.Border.BottomBorder).Value };
                Color bottomBorderColor = new Color() { Rgb = borderInfo.Border.BottomBorderColor.ToHex() };
                bottomBorder.Append(bottomBorderColor);
                border.Append(bottomBorder);

                DiagonalBorder diagonalBorder = new DiagonalBorder() { Style = borderStyleValues.Single(b => b.Key == borderInfo.Border.DiagonalBorder).Value };
                Color diagonalBorderColor = new Color() { Rgb = borderInfo.Border.DiagonalBorderColor.ToHex() };
                diagonalBorder.Append(diagonalBorderColor);
                border.Append(diagonalBorder);

                borders.Append(border);
            }



            // Cell style formats = Formats to be used by the cells and named styles
            CellStyleFormats cellStyleFormats = new CellStyleFormats() { Count = (UInt32Value)styleCount };
            // Cell formats = Any kind of formatting applied to a cell
            CellFormats cellFormats = new CellFormats() { Count = (UInt32Value)styleCount };
            foreach (var styleInfo in sharedStyles.Values)
            {
                var formatId = styleInfo.StyleId;
                var numberFormatId = styleInfo.NumberFormatId;
                var fontId = styleInfo.FontId;
                var fillId = styleInfo.FillId;
                var borderId = styleInfo.BorderId;
                Boolean applyFill = fillPatternValues.Single(p => p.Key == styleInfo.Style.Fill.PatternType).Value == PatternValues.None;
                IXLBorder opBorder = styleInfo.Style.Border;
                Boolean applyBorder = (
                       borderStyleValues.Single(b => b.Key == opBorder.BottomBorder).Value != BorderStyleValues.None
                    || borderStyleValues.Single(b => b.Key == opBorder.DiagonalBorder).Value != BorderStyleValues.None
                    || borderStyleValues.Single(b => b.Key == opBorder.RightBorder).Value != BorderStyleValues.None
                    || borderStyleValues.Single(b => b.Key == opBorder.LeftBorder).Value != BorderStyleValues.None
                    || borderStyleValues.Single(b => b.Key == opBorder.TopBorder).Value != BorderStyleValues.None);

                CellFormat cellStyleFormat = new CellFormat() { NumberFormatId = (UInt32Value)(UInt32)numberFormatId, FontId = (UInt32Value)fontId, FillId = (UInt32Value)fillId, BorderId = (UInt32Value)borderId, ApplyNumberFormat = false, ApplyFill = applyFill, ApplyBorder = applyBorder, ApplyAlignment = false, ApplyProtection = false };
                cellStyleFormats.Append(cellStyleFormat);

                CellFormat cellFormat = new CellFormat() { NumberFormatId = (UInt32Value)(UInt32)numberFormatId, FontId = (UInt32Value)fontId, FillId = (UInt32Value)fillId, BorderId = (UInt32Value)borderId, FormatId = (UInt32Value)formatId, ApplyNumberFormat = false, ApplyFill = applyFill, ApplyBorder = applyBorder, ApplyAlignment = false, ApplyProtection = false };
                Alignment alignment = new Alignment()
                {
                    Horizontal = alignmentHorizontalValues.Single(a=>a.Key== styleInfo.Style.Alignment.Horizontal).Value,
                    Vertical = alignmentVerticalValues.Single(a=>a.Key == styleInfo.Style.Alignment.Vertical).Value,
                    Indent = (UInt32)styleInfo.Style.Alignment.Indent,
                    ReadingOrder = (UInt32)styleInfo.Style.Alignment.ReadingOrder,
                    WrapText = styleInfo.Style.Alignment.WrapText,
                    TextRotation = (UInt32)styleInfo.Style.Alignment.TextRotation,
                    ShrinkToFit = styleInfo.Style.Alignment.ShrinkToFit,
                    RelativeIndent = styleInfo.Style.Alignment.RelativeIndent,
                    JustifyLastLine = styleInfo.Style.Alignment.JustifyLastLine
                };
                cellFormat.Append(alignment);

                cellFormats.Append(cellFormat);
            }



            // Cell styles = Named styles
            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)1U };
            var defaultFormatId = sharedStyles.Values.Where(s => s.Style.ToString() == DefaultStyle.ToString()).Single().StyleId;
            CellStyle cellStyle1 = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)defaultFormatId, BuiltinId = (UInt32Value)0U };
            cellStyles1.Append(cellStyle1);

            DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium9", DefaultPivotStyle = "PivotStyleLight16" };

            stylesheet1.Append(numberingFormats);
            stylesheet1.Append(fonts);
            stylesheet1.Append(fills);
            stylesheet1.Append(borders);
            stylesheet1.Append(cellStyleFormats);
            stylesheet1.Append(cellFormats);
            stylesheet1.Append(cellStyles1);
            stylesheet1.Append(differentialFormats1);
            stylesheet1.Append(tableStyles1);

            workbookStylesPart.Stylesheet = stylesheet1;
        }

        private void GenerateWorksheetPartContent(WorksheetPart worksheetPart, XLWorksheet xlWorksheet)
        {
            Worksheet worksheet = new Worksheet();
            worksheet.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            SheetProperties sheetProperties = new SheetProperties() { CodeName = xlWorksheet.Name.RemoveSpecialCharacters() };
            OutlineProperties outlineProperties = new OutlineProperties() { 
                SummaryBelow = (xlWorksheet.Outline.SummaryVLocation == XLOutlineSummaryVLocation.Bottom),
                SummaryRight = (xlWorksheet.Outline.SummaryHLocation == XLOutlineSummaryHLocation.Right)
            };
            sheetProperties.Append(outlineProperties);

            if (xlWorksheet.PageSetup.PagesTall > 0 || xlWorksheet.PageSetup.PagesWide > 0)
            {
                PageSetupProperties pageSetupProperties = new PageSetupProperties() { FitToPage = true };
                sheetProperties.Append(pageSetupProperties);
            }
            

            UInt32 maxColumn = 0;
            UInt32 maxRow = 0;

            String sheetDimensionReference = "A1";
            if (xlWorksheet.Internals.CellsCollection.Count > 0)
            {
                maxColumn = (UInt32)xlWorksheet.Internals.CellsCollection.Select(c => c.Key.ColumnNumber).Max();
                maxRow = (UInt32)xlWorksheet.Internals.CellsCollection.Select(c => c.Key.RowNumber).Max();
                sheetDimensionReference = "A1:" + new XLAddress((Int32)maxRow, (Int32)maxColumn).ToString();
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

            SheetDimension sheetDimension = new SheetDimension() { Reference = sheetDimensionReference };

            Boolean tabSelected = xlWorksheet.Name == Worksheets.Worksheet(0).Name;

            SheetViews sheetViews = new SheetViews();
            SheetView sheetView = new SheetView() { TabSelected = tabSelected, WorkbookViewId = (UInt32Value)0U };

            sheetViews.Append(sheetView);

            var maxOutlineColumn = 0;
            if (xlWorksheet.Columns().Count() > 0)
                maxOutlineColumn = xlWorksheet.Columns().Cast<XLColumn>().Max(c => c.OutlineLevel);

            var maxOutlineRow = 0;
            if (xlWorksheet.Rows().Count() > 0)
                maxOutlineRow = xlWorksheet.Rows().Cast<XLRow>().Max(c => c.OutlineLevel);

            SheetFormatProperties sheetFormatProperties3 = new SheetFormatProperties() { DefaultRowHeight = xlWorksheet.RowHeight, DefaultColumnWidth = xlWorksheet.ColumnWidth , CustomHeight = true };
            if (maxOutlineColumn > 0)
                sheetFormatProperties3.OutlineLevelColumn = (byte)maxOutlineColumn;
            if (maxOutlineRow > 0)
                sheetFormatProperties3.OutlineLevelRow = (byte)maxOutlineRow;

            Columns columns = new Columns();

            
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
                Column column = new Column()
                {
                    Min = 1,
                    Max = (UInt32Value)(UInt32)(minInColumnsCollection - 1),
                    Style = sharedStyles[xlWorksheet.Style.ToString()].StyleId,
                    Width = xlWorksheet.ColumnWidth,
                    CustomWidth = true
                };
                columns.Append(column);
            }

            for(var co = minInColumnsCollection; co <= maxInColumnsCollection; co++)
            {
                UInt32 styleId;
                Double columnWidth;
                Boolean isHidden = false;
                Boolean collapsed = false;
                Int32 outlineLevel = 0;
                if (xlWorksheet.Internals.ColumnsCollection.ContainsKey(co))
                {
                    styleId = sharedStyles[xlWorksheet.Internals.ColumnsCollection[co].Style.ToString()].StyleId;
                    columnWidth = xlWorksheet.Internals.ColumnsCollection[co].Width;
                    isHidden = xlWorksheet.Internals.ColumnsCollection[co].IsHidden;
                    collapsed = xlWorksheet.Internals.ColumnsCollection[co].Collapsed;
                    outlineLevel = xlWorksheet.Internals.ColumnsCollection[co].OutlineLevel;
                }
                else
                {
                    styleId = sharedStyles[xlWorksheet.Style.ToString()].StyleId;
                    columnWidth = xlWorksheet.ColumnWidth;
                }

                Column column = new Column()
                {
                    Min = (UInt32Value)(UInt32)co,
                    Max = (UInt32Value)(UInt32)co,
                    Style = styleId,
                    Width = columnWidth,
                    CustomWidth = true
                };
                if (isHidden) column.Hidden = true;
                if (collapsed) column.Collapsed = true;
                if (outlineLevel > 0) column.OutlineLevel = (byte)outlineLevel;
                columns.Append(column);
            }

            if (maxInColumnsCollection < XLWorksheet.MaxNumberOfColumns)
            {
                Column column = new Column()
                {
                    Min = (UInt32Value)(UInt32)(maxInColumnsCollection + 1),
                    Max = (UInt32Value)(UInt32)(XLWorksheet.MaxNumberOfColumns),
                    Style = sharedStyles[xlWorksheet.Style.ToString()].StyleId,
                    Width = xlWorksheet.ColumnWidth,
                    CustomWidth = true
                };
                columns.Append(column);
            }

            SheetData sheetData = new SheetData();

            var rowsFromCells = xlWorksheet.Internals.CellsCollection.Where(c => c.Key.ColumnNumber > 0 && c.Key.RowNumber > 0).Select(c => c.Key.RowNumber).Distinct();
            var rowsFromCollection = xlWorksheet.Internals.RowsCollection.Keys;
            var allRows = rowsFromCells.ToList();
            allRows.AddRange(rowsFromCollection);
            var distinctRows = allRows.Distinct();
            foreach (var distinctRow in distinctRows.OrderBy(r => r))
            {
                Row row = new Row() { RowIndex = (UInt32Value)(UInt32)distinctRow };
                if (maxColumn > 0)
                    row.Spans = new ListValue<StringValue>() { InnerText = "1:" + maxColumn.ToString() };

                if (xlWorksheet.Internals.RowsCollection.ContainsKey(distinctRow))
                {
                    var thisRow = xlWorksheet.Internals.RowsCollection[distinctRow];
                    var thisRowStyleString = thisRow.Style.ToString();
                    row.Height = thisRow.Height;
                    row.CustomHeight = true;
                    row.StyleIndex = sharedStyles[thisRowStyleString].StyleId;
                    row.CustomFormat = true;
                    if (thisRow.IsHidden) row.Hidden = true;
                    if (thisRow.Collapsed) row.Collapsed = true;
                    if (thisRow.OutlineLevel > 0) row.OutlineLevel = (byte)thisRow.OutlineLevel;
                }
                else
                {
                    row.Height = xlWorksheet.RowHeight;
                    row.CustomHeight = true;
                    row.Hidden = false;
                }

                foreach (var opCell in xlWorksheet.Internals.CellsCollection
                    .Where(c => c.Key.RowNumber == distinctRow)
                    .OrderBy(c => c.Key)
                    .Select(c => c))
                {
                    var styleId = sharedStyles[opCell.Value.Style.ToString()].StyleId;
                    Cell cell;
                    var dataType = opCell.Value.DataType;
                    var cellReference = opCell.Key.ToString();
                    if (!String.IsNullOrWhiteSpace(opCell.Value.FormulaA1))
                    {
                        cell = new Cell() { CellReference = cellReference, StyleIndex = styleId };
                        cell.Append(new CellFormula(opCell.Value.FormulaA1));
                    }
                    else
                    {
                        if (opCell.Value.DataType == XLCellValues.DateTime)
                        {
                            cell = new Cell()
                            {
                                CellReference = cellReference,
                                StyleIndex = styleId
                            };
                        }
                        else if (styleId == 0)
                        {
                            cell = new Cell()
                            {
                                CellReference = cellReference,
                                DataType = GetCellValue(dataType)
                            };
                        }
                        else
                        {
                            cell = new Cell()
                            {
                                CellReference = cellReference,
                                DataType = GetCellValue(dataType),
                                StyleIndex = styleId
                            };
                        }
                        CellValue cellValue = new CellValue();
                        if (dataType == XLCellValues.Text && !String.IsNullOrWhiteSpace(opCell.Value.InnerText))
                        {
                            cellValue.Text = sharedStrings[opCell.Value.InnerText].ToString();
                        }
                        else
                        {
                            cellValue.Text = opCell.Value.InnerText;
                        }
                        cell.Append(cellValue);
                    }
                    
                    row.Append(cell);
                }
                sheetData.Append(row);
            }

            MergeCells mergeCells = null;
            if (xlWorksheet.Internals.MergedCells.Count > 0)
            {
                mergeCells = new MergeCells() { Count = (UInt32Value)(UInt32)xlWorksheet.Internals.MergedCells.Count };
                foreach (var merged in xlWorksheet.Internals.MergedCells)
                {
                    MergeCell mergeCell = new MergeCell() { Reference = merged };
                    mergeCells.Append(mergeCell);
                }
            }

            PageMargins pageMargins = new PageMargins() { 
                Left = xlWorksheet.PageSetup.Margins.Left,
                Right = xlWorksheet.PageSetup.Margins.Right,
                Top = xlWorksheet.PageSetup.Margins.Top,
                Bottom = xlWorksheet.PageSetup.Margins.Bottom,
                Header = xlWorksheet.PageSetup.Margins.Header,
                Footer = xlWorksheet.PageSetup.Margins.Footer
            };

            

            //Drawing drawing1 = new Drawing() { Id = "rId1" };

            PageSetup pageSetup1 = new PageSetup() { 
                Orientation = pageOrientationValues.Single(p=>p.Key == xlWorksheet.PageSetup.PageOrientation).Value,
                Id = "rId" + RelId.GetNext(RelType.Worksheet),
                PaperSize = (UInt32Value)(UInt32)xlWorksheet.PageSetup.PaperSize,
                BlackAndWhite = xlWorksheet.PageSetup.BlackAndWhite,
                Draft = xlWorksheet.PageSetup.DraftQuality,
                PageOrder = pageOrderValues.Single(p=>p.Key == xlWorksheet.PageSetup.PageOrder).Value,
                CellComments = showCommentsValues.Single(s=>s.Key == xlWorksheet.PageSetup.ShowComments).Value,
                Errors = printErrorValues.Single(p=>p.Key == xlWorksheet.PageSetup.PrintErrorValue).Value
            };

            if (xlWorksheet.PageSetup.FirstPageNumber > 0)
            {
                pageSetup1.FirstPageNumber = (UInt32Value)(UInt32)xlWorksheet.PageSetup.FirstPageNumber;
                pageSetup1.UseFirstPageNumber = true;
            }

            if (xlWorksheet.PageSetup.HorizontalDpi > 0)
                pageSetup1.HorizontalDpi = (UInt32Value)(UInt32)xlWorksheet.PageSetup.HorizontalDpi;

            if (xlWorksheet.PageSetup.VerticalDpi > 0)
                pageSetup1.VerticalDpi = (UInt32Value)(UInt32)xlWorksheet.PageSetup.VerticalDpi;

            if (xlWorksheet.PageSetup.Scale > 0)
            {
                pageSetup1.Scale = (UInt32Value)(UInt32)xlWorksheet.PageSetup.Scale;
            }
            else
            {
                if (xlWorksheet.PageSetup.PagesWide > 0)
                    pageSetup1.FitToWidth = (UInt32Value)(UInt32)xlWorksheet.PageSetup.PagesWide;
                if (xlWorksheet.PageSetup.PagesTall > 0)
                    pageSetup1.FitToHeight = (UInt32Value)(UInt32)xlWorksheet.PageSetup.PagesTall;
            }

            PrintOptions printOptions = new PrintOptions()
            {
                HorizontalCentered = xlWorksheet.PageSetup.CenterHorizontally,
                VerticalCentered = xlWorksheet.PageSetup.CenterVertically,
                Headings = xlWorksheet.PageSetup.ShowRowAndColumnHeadings,
                GridLines = xlWorksheet.PageSetup.ShowGridlines
            };

            HeaderFooter headerFooter = new HeaderFooter();
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

            //var firstHeaderText = "&L" + xlWorksheet.PageSetup.Header.Left.GetText(XLHFOccurrence.FirstPage) + "&C" + xlWorksheet.PageSetup.Header.Center.GetText(XLHFOccurrence.FirstPage) + "&R" + xlWorksheet.PageSetup.Header.Right.GetText(XLHFOccurrence.FirstPage) + "";

            FirstHeader firstHeader = new FirstHeader(xlWorksheet.PageSetup.Header.GetText(XLHFOccurrence.FirstPage));
            headerFooter.Append(firstHeader);
            FirstFooter firstFooter = new FirstFooter(xlWorksheet.PageSetup.Footer.GetText(XLHFOccurrence.FirstPage));
            headerFooter.Append(firstFooter);

            RowBreaks rowBreaks = null;
            var rowBreakCount = xlWorksheet.PageSetup.RowBreaks.Count;
            if (rowBreakCount > 0)
            {
                rowBreaks = new RowBreaks() { Count = (UInt32Value)(UInt32)rowBreakCount, ManualBreakCount = (UInt32)rowBreakCount };
                foreach (var rb in xlWorksheet.PageSetup.RowBreaks)
                {
                    Break break1 = new Break() { Id = (UInt32Value)(UInt32)rb, Max = (UInt32Value)(UInt32)xlWorksheet.RangeAddress.LastAddress.RowNumber, ManualPageBreak = true };
                    rowBreaks.Append(break1);
                }
               
            }

            ColumnBreaks columnBreaks = null;
            var columnBreakCount = xlWorksheet.PageSetup.ColumnBreaks.Count;
            if (columnBreakCount > 0)
            {
                columnBreaks = new ColumnBreaks() { Count = (UInt32Value)(UInt32)columnBreakCount, ManualBreakCount = (UInt32Value)(UInt32)columnBreakCount };
                foreach (var cb in xlWorksheet.PageSetup.ColumnBreaks)
                {
                    Break break1 = new Break() { Id = (UInt32Value)(UInt32)cb, Max = (UInt32Value)(UInt32)xlWorksheet.RangeAddress.LastAddress.ColumnNumber, ManualPageBreak = true };
                    columnBreaks.Append(break1);
                }
            }
            
            worksheet.Append(sheetProperties);
            worksheet.Append(sheetDimension);
            worksheet.Append(sheetViews);
            worksheet.Append(sheetFormatProperties3);
            if (columns != null) worksheet.Append(columns);
            worksheet.Append(sheetData);
            if (mergeCells != null) worksheet.Append(mergeCells);
            worksheet.Append(printOptions);
            worksheet.Append(pageMargins);
            worksheet.Append(pageSetup1);
            if (headerFooter.Any(hf=>hf.InnerText.Length > 0))
                worksheet.Append(headerFooter);
            if (rowBreaks != null) worksheet.Append(rowBreaks);
            if (columnBreaks != null) worksheet.Append(columnBreaks);
            //worksheet.Append(drawing1);

            worksheetPart.Worksheet = worksheet;
        }

        private void GenerateCalculationChainPartContent(WorkbookPart workbookPart, String rId)
        {
            Boolean foundOne = false;
            CalculationChain calculationChain = new CalculationChain();
            Int32 sheetId = 0;
            foreach (var worksheet in Worksheets.Cast<XLWorksheet>())
            {
                sheetId++;
                foreach (var c in worksheet.Internals.CellsCollection.Values.Where(c => !String.IsNullOrWhiteSpace(c.FormulaA1)))
                {
                    CalculationCell calculationCell = new CalculationCell() { CellReference = c.Address.ToString(), SheetId = sheetId };
                    calculationChain.Append(calculationCell);
                    if (!foundOne) foundOne = true;
                }
            }
            if (foundOne)
            {
                CalculationChainPart calculationChainPart = workbookPart.AddNewPart<CalculationChainPart>(rId);
                calculationChainPart.CalculationChain = calculationChain;
            }
        }

        private void GenerateThemePartContent(ThemePart themePart)
        {
            A.Theme theme1 = new A.Theme() { Name = "Office Theme" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements1 = new A.ThemeElements();

            A.ColorScheme colorScheme1 = new A.ColorScheme() { Name = "Office" };

            A.Dark1Color dark1Color1 = new A.Dark1Color();
            A.SystemColor systemColor1 = new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color1.Append(systemColor1);

            A.Light1Color light1Color1 = new A.Light1Color();
            A.SystemColor systemColor2 = new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color1.Append(systemColor2);

            A.Dark2Color dark2Color1 = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "1F497D" };

            dark2Color1.Append(rgbColorModelHex1);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "EEECE1" };

            light2Color1.Append(rgbColorModelHex2);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "4F81BD" };

            accent1Color1.Append(rgbColorModelHex3);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "C0504D" };

            accent2Color1.Append(rgbColorModelHex4);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "9BBB59" };

            accent3Color1.Append(rgbColorModelHex5);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "8064A2" };

            accent4Color1.Append(rgbColorModelHex6);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "4BACC6" };

            accent5Color1.Append(rgbColorModelHex7);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "F79646" };

            accent6Color1.Append(rgbColorModelHex8);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "0000FF" };

            hyperlink1.Append(rgbColorModelHex9);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "800080" };

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
    }
}