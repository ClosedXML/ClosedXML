using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel.Style;
using System.IO;
using System.Drawing;

namespace ClosedXML.Excel
{
    public partial class XLWorkbook: IXLWorkbook
    {
        public XLWorkbook()
        {
            Worksheets = new XLWorksheets();
        }
        #region IXLWorkbook Members

        public IXLWorksheets Worksheets { get; private set; }

        /// <summary>
        /// Gets the file name of the workbook.
        /// </summary>
        public String Name { get; private set; }

        /// <summary>
        /// Gets the file name of the workbook including its full directory.
        /// </summary>
        public String FullName { get; private set; }

        public void SaveAs(String file, Boolean overwrite = false)
        {
            if (overwrite && File.Exists(file)) File.Delete(file);

            // For maintainability reasons the XLWorkbook class was divided into two files.
            // The method CreatePackage can be located in the file XLWorkbook_Save.cs   
            CreatePackage(file);
        }

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

                        Border = new XLBorder(null)
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
                        NumberFormat = new XLNumberFormat(null) { NumberFormatId = 0 },
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
                }
                return defaultStyle;
            }
        }

        #endregion
    }
}
