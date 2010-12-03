using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.IO;
using System.Drawing;

namespace ClosedXML.Excel
{
    public enum XLCalculateMode { Auto, AutoNoTable, Manual, Default };
    public enum XLReferenceStyle { R1C1, A1, Default };
    public partial class XLWorkbook
    {
        public XLWorkbook()
        {
            DefaultRowHeight = 15;
            DefaultColumnWidth = 9.140625;
            Worksheets = new XLWorksheets(this);
            NamedRanges = new XLNamedRanges(this);
            PopulateEnums();
            Style = DefaultStyle;
            RowHeight = DefaultRowHeight;
            ColumnWidth = DefaultColumnWidth;
            PageOptions = DefaultPageOptions;
            Outline = DefaultOutline;
            Properties = new XLWorkbookProperties();
            CalculateMode = XLCalculateMode.Default;
            ReferenceStyle = XLReferenceStyle.Default;
        }

        private String originalFile;
        public XLWorkbook(String file): this()
        {
            originalFile = file;
            Load(file);
        }

        #region IXLWorkbook Members

        public IXLWorksheets Worksheets { get; private set; }
        public IXLNamedRanges NamedRanges { get; private set; }

        /// <summary>
        /// Gets the file name of the workbook.
        /// </summary>
        public String Name { get; private set; }

        /// <summary>
        /// Gets the file name of the workbook including its full directory.
        /// </summary>
        public String FullName { get; private set; }

        public void Save()
        {
            if (originalFile == null)
                throw new Exception("This is a new file, please use one of the following methods: SaveAs, MergeInto, or SaveChangesTo");

            MergeInto(originalFile);
        }

        public void SaveAs(String file)
        {
            if (originalFile == null)
                File.Delete(file);
            else if (originalFile.Trim().ToLower() != file.Trim().ToLower())
                File.Copy(originalFile, file, true);

            CreatePackage(file);
        }

        public void MergeInto(String file)
        {
            CreatePackage(file);
        }

        public void SaveChangesTo(String file)
        {
            if (File.Exists(file)) File.Delete(file);
            CreatePackage(file);
        }

        public IXLStyle Style { get; set; }
        public Double RowHeight { get; set; }
        public Double ColumnWidth { get; set; }
        public IXLPageSetup PageOptions { get; set; }
        public IXLOutline Outline { get; set; }
        public XLWorkbookProperties Properties { get; set; }
        public XLCalculateMode CalculateMode { get; set; }
        public XLReferenceStyle ReferenceStyle { get; set; }

        #endregion

        #region Static

        public static IXLStyle DefaultStyle
        {
            get
            {
                var defaultStyle = new XLStyle(null, null)
                {
                    Font = new XLFont(null, null)
                    {
                        Bold = false,
                        Italic = false,
                        Underline = XLFontUnderlineValues.None,
                        Strikethrough = false,
                        VerticalAlignment = XLFontVerticalTextAlignmentValues.Baseline,
                        FontSize = 11,
                        FontColor = Color.FromArgb(0, 0, 0),
                        FontName = "Calibri",
                        FontFamilyNumbering = XLFontFamilyNumberingValues.Swiss
                    },

                    Fill = new XLFill(null)
                   {
                       BackgroundColor = Color.FromArgb(255, 255, 255),
                       PatternType = XLFillPatternValues.None,
                       PatternColor = Color.FromArgb(255, 255, 255)
                   },

                    Border = new XLBorder(null, null)
                        {
                            BottomBorder = XLBorderStyleValues.None,
                            DiagonalBorder = XLBorderStyleValues.None,
                            DiagonalDown = false,
                            DiagonalUp = false,
                            LeftBorder = XLBorderStyleValues.None,
                            RightBorder = XLBorderStyleValues.None,
                            TopBorder = XLBorderStyleValues.None,
                            BottomBorderColor = Color.Black,
                            DiagonalBorderColor = Color.Black,
                            LeftBorderColor = Color.Black,
                            RightBorderColor = Color.Black,
                            TopBorderColor = Color.Black
                        },
                    NumberFormat = new XLNumberFormat(null, null) { NumberFormatId = 0 },
                    Alignment = new XLAlignment(null)
                        {
                            Horizontal = XLAlignmentHorizontalValues.General,
                            Indent = 0,
                            JustifyLastLine = false,
                            ReadingOrder = XLAlignmentReadingOrderValues.ContextDependent,
                            RelativeIndent = 0,
                            ShrinkToFit = false,
                            TextRotation = 0,
                            Vertical = XLAlignmentVerticalValues.Bottom,
                            WrapText = false
                        }
                };
                return defaultStyle;
            }
        }

        public static Double DefaultRowHeight { get; private set; }
        public static Double DefaultColumnWidth { get; private set; }

        public static IXLPageSetup DefaultPageOptions
        {
            get
            {
                var defaultPageOptions = new XLPageSetup(null, null)
                {
                    PageOrientation = XLPageOrientation.Default,
                    Scale = 100,
                    PaperSize = XLPaperSize.LetterPaper,
                    Margins = new XLMargins()
                    {
                        Top = 0.75,
                        Bottom = 0.75,
                        Left = 0.75,
                        Right = 0.75,
                        Header = 0.75,
                        Footer = 0.75
                    },
                    ScaleHFWithDocument = true,
                    AlignHFWithMargins = true,
                    PrintErrorValue = XLPrintErrorValues.Displayed,
                    ShowComments = XLShowCommentsValues.None
                };
                return defaultPageOptions;
            }
        }

        public static IXLOutline DefaultOutline
        {
            get
            {
                return new XLOutline(null) { 
                    SummaryHLocation = XLOutlineSummaryHLocation.Right, 
                    SummaryVLocation= XLOutlineSummaryVLocation.Bottom };
            }
        }

        public static IXLFont GetXLFont()
        {
            return new XLFont();
        }

        #endregion
    }
}
