using System;

namespace ClosedXML.Excel
{
    internal class XLComment : XLFormattedText<IXLComment>, IXLComment
    {
        private XLCell _cell;

        public XLComment(XLCell cell, IXLFontBase defaultFont = null, int? shapeId = null)
            : base(defaultFont ?? XLFont.DefaultCommentFont)
        {
            Initialize(cell, shapeId: shapeId);
        }

        public XLComment(XLCell cell, XLFormattedText<IXLComment> defaultComment, IXLFontBase defaultFont, IXLDrawingStyle style)
            : base(defaultComment, defaultFont)
        {
            Initialize(cell, style);
        }

        public XLComment(XLCell cell, String text, IXLFontBase defaultFont)
            : base(text, defaultFont)
        {
            Initialize(cell);
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

        public IXLComment SetVisible(Boolean hidden)
        {
            Visible = hidden;
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
