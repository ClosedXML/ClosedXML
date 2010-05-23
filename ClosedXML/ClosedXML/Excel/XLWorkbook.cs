using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using ClosedXML.Excel.Style;

namespace ClosedXML.Excel
{
    /// <summary>
    /// Represents an Excel Workbook file
    /// </summary>
    public partial class XLWorkbook
    {
        #region Properties

        /// <summary>
        /// Gets an instance of the <see cref="XLWorksheets"/> class.
        /// It allows you to add, access, and remove worksheets from the workbook.
        /// </summary>
        public XLWorksheets Worksheets { get; private set; }

        /// <summary>
        /// Gets the file name of the workbook.
        /// </summary>
        public String Name { get; private set; }

        /// <summary>
        /// Gets the file name of the workbook including its full directory.
        /// </summary>
        public String FullName { get; private set; }

        /// <summary>
        /// Gets the default style for new cells in this workbook.
        /// </summary>
        public XLStyle WorkbookStyle { get; private set; }

        #endregion

        #region Static

        private static XLStyle defaultStyle;
        /// <summary>
        /// Gets the default style for new workbooks.
        /// </summary>
        public static XLStyle DefaultStyle
        {
            get
            {
                if (defaultStyle == null)
                {
                    defaultStyle = new XLStyle(null, null)
                    {
                        Font = new XLFont(null, null)
                        {
                            Bold = false,
                            Italic = false,
                            Underline = XLFontUnderlineValues.None,
                            Strikethrough = false,
                            VerticalAlignment = XLFontVerticalTextAlignmentValues.Baseline,
                            FontSize = 11,
                            Color = "FF000000",
                            FontName = "Calibri",
                            FontFamilyNumbering = 2
                        },
                        Fill = new XLFill(null, null)
                        {
                            BackgroundColor = "FFFFFFFF",
                            PatternType = XLFillPatternValues.None
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
                            BottomBorderColor = "000000",
                            DiagonalBorderColor = "000000",
                            LeftBorderColor = "000000",
                            RightBorderColor = "000000",
                            TopBorderColor = "000000"
                        },
                        NumberFormat = new XLNumberFormat(null, null),
                        Alignment = new XLAlignment(null, null)
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
                }
                return defaultStyle;
            }
        }

        #endregion

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="XLWorkbook"/> class.
        /// </summary>
        /// <param name="file">New Excel file to be created.</param>
        public XLWorkbook(String file)
        {
            if (File.Exists(file)) File.Delete(file);
                //throw new ArgumentException("File already exists.");

            FileInfo fi = new FileInfo(file);
            this.Name = fi.Name;
            this.FullName = fi.FullName;
            Worksheets = new XLWorksheets(this);

            WorkbookStyle = new XLStyle(XLWorkbook.DefaultStyle, null);
        }

        #endregion

        #region Methods

        /// <summary>
        /// Saves the current workbook to disk.
        /// </summary>
        public void Save()
        {
            // For maintainability reasons the XLWorkbook class was divided into two files.
            // The method CreatePackage can be located in the file XLWorkbook_Save.cs
            CreatePackage(this.FullName);
        }

        #endregion

    }
}
