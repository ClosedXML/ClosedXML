using System;

namespace ClosedXML.Excel
{
    internal class XLComment : XLFormattedText<IXLComment>, IXLComment
    {
        private XLCell _cell;

        public XLComment(XLCell cell, IXLFontBase defaultFont)
            : base(defaultFont)
        {
            Initialize(cell);
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

        private Boolean _visible;

        public Boolean Visible
        {
            get
            {
                return _visible;
            }
            set
            {
                _visible = value;
            }
        }

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

        private void Initialize(XLCell cell)
        {
            var style = new XLDrawingStyle();

            style
                .Margins.SetLeft(0.1)
                .Margins.SetRight(0.1)
                .Margins.SetTop(0.05)
                .Margins.SetBottom(0.05)
                .Margins.SetAutomatic()
                .Size.SetHeight(59.25)
                .Size.SetWidth(19.2)
                .ColorsAndLines.SetLineColor(XLColor.Black)
                .ColorsAndLines.SetFillColor(XLColor.FromArgb(255, 255, 225))
                .ColorsAndLines.SetLineDash(XLDashStyle.Solid)
                .ColorsAndLines.SetLineStyle(XLLineStyle.Single)
                .ColorsAndLines.SetLineWeight(0.75)
                .ColorsAndLines.SetFillTransparency(1)
                .ColorsAndLines.SetLineTransparency(1)
                .Alignment.SetHorizontal(XLDrawingHorizontalAlignment.Left)
                .Alignment.SetVertical(XLDrawingVerticalAlignment.Top)
                .Alignment.SetDirection(XLDrawingTextDirection.Context)
                .Alignment.SetOrientation(XLDrawingTextOrientation.LeftToRight)
                .Properties.SetPositioning(XLDrawingAnchor.Absolute)
                .Protection.SetLocked()
                .Protection.SetLockText();

            Initialize(cell, style);
        }

        private void Initialize(XLCell cell, IXLDrawingStyle style)
        {
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
            ShapeId = cell.Worksheet.Workbook.ShapeIdManager.GetNext();
        }
    }
}
