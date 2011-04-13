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
        private enum XLLoadSource { New, File, Stream };
        private XLLoadSource loadSource = XLLoadSource.New;
        /// <summary>
        /// Creates a new Excel workbook.
        /// </summary>
        public XLWorkbook()
        {
            DefaultRowHeight = 15;
            DefaultColumnWidth = 9.140625;
            Worksheets = new XLWorksheets(this);
            NamedRanges = new XLNamedRanges(this);
            CustomProperties = new XLCustomProperties(this);
            PopulateEnums();
            Style = DefaultStyle;
            RowHeight = DefaultRowHeight;
            ColumnWidth = DefaultColumnWidth;
            PageOptions = DefaultPageOptions;
            Outline = DefaultOutline;
            Properties = new XLWorkbookProperties();
            CalculateMode = XLCalculateMode.Default;
            ReferenceStyle = XLReferenceStyle.Default;
            InitializeTheme();
        }

        private void InitializeTheme()
        {
            Theme = new XLTheme() {
                Text1 = XLColor.FromHtml("#FF000000"),
                Background1 = XLColor.FromHtml("#FFFFFFFF"),
                Text2 = XLColor.FromHtml("#FF1F497D"),
                Background2 = XLColor.FromHtml("#FFEEECE1"),
                Accent1 = XLColor.FromHtml("#FF4F81BD"),
                Accent2 = XLColor.FromHtml("#FFC0504D"),
                Accent3 = XLColor.FromHtml("#FF9BBB59"),
                Accent4 = XLColor.FromHtml("#FF8064A2"),
                Accent5 = XLColor.FromHtml("#FF4BACC6"),
                Accent6 = XLColor.FromHtml("#FFF79646"),
                Hyperlink = XLColor.FromHtml("#FF0000FF"),
                FollowedHyperlink = XLColor.FromHtml("#FF800080")
            };
        }

        internal IXLColor GetXLColor(XLThemeColor themeColor)
        {
            switch (themeColor)
            {
                case XLThemeColor.Text1: return Theme.Text1;
                case XLThemeColor.Background1: return Theme.Background1;
                case XLThemeColor.Text2: return Theme.Text2;
                case XLThemeColor.Background2: return Theme.Background2;
                case XLThemeColor.Accent1: return Theme.Accent1;
                case XLThemeColor.Accent2: return Theme.Accent2;
                case XLThemeColor.Accent3: return Theme.Accent3;
                case XLThemeColor.Accent4: return Theme.Accent4;
                case XLThemeColor.Accent5: return Theme.Accent5;
                case XLThemeColor.Accent6: return Theme.Accent6;
                default: throw new ArgumentException("Invalid theme color");
            }
        }

        private String originalFile;
        private Stream originalStream;
        /// <summary>
        /// Opens an existing workbook from a file.
        /// </summary>
        /// <param name="file">The file to open.</param>
        public XLWorkbook(String file): this()
        {
            loadSource = XLLoadSource.File;
            originalFile = file;
            Load(file);
        }

        /// <summary>
        /// Opens an existing workbook from a stream.
        /// </summary>
        /// <param name="stream">The stream to open.</param>
        public XLWorkbook(Stream stream)
            : this()
        {
            loadSource = XLLoadSource.Stream;
            originalStream = stream;
            Load(stream);
        }

        #region IXLWorkbook Members

        /// <summary>
        /// Gets an object to manipulate the worksheets.
        /// </summary>
        public IXLWorksheets Worksheets { get; private set; }
        /// <summary>
        /// Gets an object to manipulate this workbook's named ranges.
        /// </summary>
        public IXLNamedRanges NamedRanges { get; private set; }

        public IXLNamedRange NamedRange(String rangeName)
        {
            return NamedRanges.NamedRange(rangeName);
        }

        /// <summary>
        /// Gets the file name of the workbook.
        /// </summary>
        public String Name { get; private set; }

        /// <summary>
        /// Gets the file name of the workbook including its full directory.
        /// </summary>
        public String FullName { get; private set; }
        /// <summary>
        /// Gets an object to manipulate this workbook's theme.
        /// </summary>
        public IXLTheme Theme { get; private set; }

        /// <summary>
        /// Saves the current workbook.
        /// </summary>
        public void Save()
        {
            if (loadSource == XLLoadSource.New)
                throw new Exception("This is a new file, please use one of the SaveAs methods.");
            else if (loadSource == XLLoadSource.Stream)
                CreatePackage(originalStream, false);
            else
                CreatePackage(originalFile);
        }
        /// <summary>
        /// Saves the current workbook to a file.
        /// </summary>
        public void SaveAs(String file)
        {
            if (loadSource == XLLoadSource.New)
            {
                if (File.Exists(file))
                    File.Delete(file);

                CreatePackage(file);
            }
            else if (loadSource == XLLoadSource.File)
            {
                if (originalFile.Trim().ToLower() != file.Trim().ToLower())
                    File.Copy(originalFile, file, true);

                CreatePackage(file);
            }
            else if (loadSource == XLLoadSource.Stream)
            {
                originalStream.Position = 0;
                
                using (var fileStream = File.Create(file))
                {
                    CopyStream(originalStream, fileStream);
                    //fileStream.Position = 0;
                    CreatePackage(fileStream, false);
                    fileStream.Close();
                }
            }
        }
        /// <summary>
        /// Saves the current workbook to a stream.
        /// </summary>
        public void SaveAs(Stream stream)
        {
            if (loadSource == XLLoadSource.New)
            {
                CreatePackage(stream, true);
            }
            else if (loadSource == XLLoadSource.File)
            {
                using (var fileStream = new FileStream(originalFile, FileMode.Open))
                {
                    CopyStream(fileStream, stream);
                    fileStream.Close();
                }
                CreatePackage(stream, false);
            }
            else if (loadSource == XLLoadSource.Stream)
            {
                originalStream.Position = 0;
                if (originalStream != stream)
                    CopyStream(originalStream, stream);

                CreatePackage(stream, false);
            }
        }

        internal void CopyStream(Stream input, Stream output)
        {
            byte[] buffer = new byte[8 * 1024];
            int len;
            while ((len = input.Read(buffer, 0, buffer.Length)) > 0)
            {
                output.Write(buffer, 0, len);
            }  
        }


        /// <summary>
        /// Gets or sets the default style for the workbook.
        /// <para>All new worksheets will use this style.</para>
        /// </summary>
        public IXLStyle Style { get; set; }
        /// <summary>
        /// Gets or sets the default row height for the workbook.
        /// <para>All new worksheets will use this row height.</para>
        /// </summary>
        public Double RowHeight { get; set; }
        /// <summary>
        /// Gets or sets the default column width for the workbook.
        /// <para>All new worksheets will use this column width.</para>
        /// </summary>
        public Double ColumnWidth { get; set; }
        /// <summary>
        /// Gets or sets the default page options for the workbook.
        /// <para>All new worksheets will use these page options.</para>
        /// </summary>
        public IXLPageSetup PageOptions { get; set; }
        /// <summary>
        /// Gets or sets the default outline options for the workbook.
        /// <para>All new worksheets will use these outline options.</para>
        /// </summary>
        public IXLOutline Outline { get; set; }
        /// <summary>
        /// Gets or sets the workbook's properties.
        /// </summary>
        public XLWorkbookProperties Properties { get; set; }
        /// <summary>
        /// Gets or sets the workbook's calculation mode.
        /// </summary>
        public XLCalculateMode CalculateMode { get; set; }
        /// <summary>
        /// Gets or sets the workbook's reference style.
        /// </summary>
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
                        FontColor = XLColor.FromArgb(0, 0, 0),
                        FontName = "Calibri",
                        FontFamilyNumbering = XLFontFamilyNumberingValues.Swiss
                    },

                    Fill = new XLFill(null)
                   {
                       BackgroundColor = XLColor.FromIndex(64),
                       PatternType = XLFillPatternValues.None,
                       PatternColor = XLColor.FromIndex(64)
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
                            BottomBorderColor = XLColor.Black,
                            DiagonalBorderColor = XLColor.Black,
                            LeftBorderColor = XLColor.Black,
                            RightBorderColor = XLColor.Black,
                            TopBorderColor = XLColor.Black
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
                            },
                        Protection = new XLProtection(null)
                            {
                                Locked = true,
                                Hidden = false
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

        public IXLWorksheet Worksheet(String name)
        {
            return Worksheets.Worksheet(name);
        }
        public IXLWorksheet Worksheet(Int32 position)
        {
            return Worksheets.Worksheet(position);
        }

        public HashSet<String> GetSharedStrings()
        {
            HashSet<String> modifiedStrings = new HashSet<String>();
            foreach (var w in Worksheets.Cast<XLWorksheet>())
            {
                foreach (var c in w.Internals.CellsCollection.Values)
                {
                    if (
                        c.DataType == XLCellValues.Text
                        && !StringExtensions.IsNullOrWhiteSpace(c.InnerText)
                        && c.ShareString
                        && !modifiedStrings.Contains(c.Value.ToString())
                        )
                    {
                        modifiedStrings.Add(c.Value.ToString());
                    }
                }
            }
            return modifiedStrings;
        }

        public IXLCustomProperty CustomProperty(String name)
        {
            return CustomProperties.CustomProperty(name);
        }

        public IXLCustomProperties CustomProperties { get; private set; }
    }
}
