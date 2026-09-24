#nullable disable warnings

using System;
using System.Diagnostics;
using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel
{
    internal class XLComment : XLFormattedText<IXLComment>, IXLComment
    {
        private XLCell _cell;

        private static XLFontFormatValue DefaultCommentFont
        {
            get
            {
                // MS Excel uses Tahoma 9 Swiss no matter what current style font
                var defaultCommentFont = XLFontFormatValue.Default with
                {
                    Name = "Tahoma",
                    Size = XLFontSize.FromPoints(9),
                    Family = XLFontFamilyNumberingValues.Swiss,
                    Color = XLColor.Black
                };

                return defaultCommentFont;
            }
        }

        private XLComment(XLFontFormatValue defaultFont, XLWorkbookStyles styles, XLFontFormatValue? phoneticsFont)
            : base(defaultFont, styles)
        {
            if (phoneticsFont is not null)
            {
                Debug.Assert(styles.Fonts.ContainsValue(phoneticsFont));
                Phonetics = new XLPhonetics(phoneticsFont, defaultFont, styles, OnContentChanged);
            }
        }

        #region IXLComment Members

        public String Author { get; set; }

        public IXLComment SetAuthor(String value)
        {
            Author = value;
            return this;
        }

        public IXLRichString AddSignature()
        {
            AddText(Author + ":").SetBold();
            return AddText(Environment.NewLine);
        }

        public void Delete()
        {
            _cell.DeleteComment();
        }

        #endregion IXLComment Members

        #region IXLDrawing

        public String Name { get; set; }
        public String Description { get; set; }
        public XLDrawingAnchor Anchor { get; set; }
        public Boolean HorizontalFlip { get; set; }
        public Boolean VerticalFlip { get; set; }
        public Int32 Rotation { get; set; }
        public Int32 ExtentLength { get; set; }
        public Int32 ExtentWidth { get; set; }
        public Int32 ShapeId { get; internal set; }
        public Boolean Visible { get; set; }

        public IXLComment SetVisible()
        {
            Visible = true;
            return Container;
        }

        public IXLComment SetVisible(Boolean visible)
        {
            Visible = visible;
            return Container;
        }

        public IXLDrawingPosition Position { get; private set; }

        public Int32 ZOrder { get; set; }

        public IXLComment SetZOrder(Int32 zOrder)
        {
            ZOrder = zOrder;
            return Container;
        }

        public IXLDrawingStyle Style { get; private set; }

        public IXLComment SetName(String name)
        {
            Name = name;
            return Container;
        }

        public IXLComment SetDescription(String description)
        {
            Description = description;
            return Container;
        }

        public IXLComment SetHorizontalFlip()
        {
            HorizontalFlip = true;
            return Container;
        }

        public IXLComment SetHorizontalFlip(Boolean horizontalFlip)
        {
            HorizontalFlip = horizontalFlip;
            return Container;
        }

        public IXLComment SetVerticalFlip()
        {
            VerticalFlip = true;
            return Container;
        }

        public IXLComment SetVerticalFlip(Boolean verticalFlip)
        {
            VerticalFlip = verticalFlip;
            return Container;
        }

        public IXLComment SetRotation(Int32 rotation)
        {
            Rotation = rotation;
            return Container;
        }

        public IXLComment SetExtentLength(Int32 extentLength)
        {
            ExtentLength = extentLength;
            return Container;
        }

        public IXLComment SetExtentWidth(Int32 extentWidth)
        {
            ExtentWidth = extentWidth;
            return Container;
        }

        #endregion IXLDrawing

        internal static XLComment Create(XLCell cell, int? shapeId)
        {
            var styles = cell.Worksheet.Workbook.Styles;
            var defaultFont = styles.RegisterFontFormat(DefaultCommentFont);
            var comment = new XLComment(defaultFont, styles, null);
            comment.Initialize(cell, shapeId: shapeId);
            return comment;
        }

        internal static XLComment CreateAsCopy(XLCell targetCell, XLCell sourceCell, XLComment originalComment)
        {
            // source cell could be from different workbook, so register formats
            var styles = targetCell.Worksheet.Workbook.Styles;
            var defaultFont = styles.RegisterFontFormat(sourceCell.GetFormat().Font);
            var phoneticsFont = originalComment.HasPhonetics
                ? styles.RegisterFontFormat(XLFontFormatValue.FromFontBase(originalComment.Phonetics, styles))
                : null;
            var comment = new XLComment(defaultFont, styles, phoneticsFont);

            foreach (XLRichString rt in originalComment)
                comment.AddText(rt.Text, rt);

            comment.Initialize(targetCell, originalComment.Style);
            return comment;
        }

        private void Initialize(XLCell cell, IXLDrawingStyle style = null, int? shapeId = null)
        {
            style = style ?? XLDrawingStyle.DefaultCommentStyle;
            shapeId = shapeId ?? cell.Worksheet.Workbook.ShapeIdManager.GetNext();

            Author = cell.Worksheet.Author;
            Container = this;
            Anchor = XLDrawingAnchor.MoveAndSizeWithCells;
            Style = new XLDrawingStyle();
            Int32 previousRowNumber = cell.Address.RowNumber;
            Double previousRowOffset = 0;

            if (previousRowNumber > 1)
            {
                previousRowNumber--;

                if (cell.Worksheet.Internals.RowsCollection.TryGetValue(previousRowNumber, out XLRow previousRow))
                    previousRowOffset = Math.Max(0, previousRow.Height - 7);
                else
                    previousRowOffset = Math.Max(0, cell.Worksheet.RowHeight - 7);
            }

            Position = new XLDrawingPosition
            {
                Column = cell.Address.ColumnNumber + 1,
                ColumnOffset = 2,
                Row = previousRowNumber,
                RowOffset = previousRowOffset
            };

            ZOrder = cell.Worksheet.ZOrder++;
            Style
                .Margins.SetLeft(style.Margins.Left)
                .Margins.SetRight(style.Margins.Right)
                .Margins.SetTop(style.Margins.Top)
                .Margins.SetBottom(style.Margins.Bottom)
                .Margins.SetAutomatic(style.Margins.Automatic)
                .Size.SetHeight(style.Size.Height)
                .Size.SetWidth(style.Size.Width)
                .ColorsAndLines.SetLineColor(style.ColorsAndLines.LineColor)
                .ColorsAndLines.SetFillColor(style.ColorsAndLines.FillColor)
                .ColorsAndLines.SetLineDash(style.ColorsAndLines.LineDash)
                .ColorsAndLines.SetLineStyle(style.ColorsAndLines.LineStyle)
                .ColorsAndLines.SetLineWeight(style.ColorsAndLines.LineWeight)
                .ColorsAndLines.SetFillTransparency(style.ColorsAndLines.FillTransparency)
                .ColorsAndLines.SetLineTransparency(style.ColorsAndLines.LineTransparency)
                .Alignment.SetHorizontal(style.Alignment.Horizontal)
                .Alignment.SetVertical(style.Alignment.Vertical)
                .Alignment.SetDirection(style.Alignment.Direction)
                .Alignment.SetOrientation(style.Alignment.Orientation)
                .Alignment.SetAutomaticSize(style.Alignment.AutomaticSize)
                .Properties.SetPositioning(style.Properties.Positioning)
                .Protection.SetLocked(style.Protection.Locked)
                .Protection.SetLockText(style.Protection.LockText);

            _cell = cell;
            ShapeId = shapeId.Value;
        }
    }
}
