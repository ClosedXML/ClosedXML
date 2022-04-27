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

        public XLComment(XLCell cell, string text, IXLFontBase defaultFont)
            : base(text, defaultFont)
        {
            Initialize(cell);
        }

        #region IXLComment Members

        public string Author { get; set; }

        public IXLComment SetAuthor(string value)
        {
            Author = value;
            return this;
        }

        public IXLRichString AddSignature()
        {
            AddText(Author + ":").SetBold();
            return AddText(XLConstants.NewLine);
        }

        public void Delete()
        {
            _cell.DeleteComment();
        }

        #endregion IXLComment Members

        #region IXLDrawing

        public string Name { get; set; }
        public string Description { get; set; }
        public XLDrawingAnchor Anchor { get; set; }
        public bool HorizontalFlip { get; set; }
        public bool VerticalFlip { get; set; }
        public int Rotation { get; set; }
        public int ExtentLength { get; set; }
        public int ExtentWidth { get; set; }
        public int ShapeId { get; internal set; }
        public bool Visible { get; set; }

        public IXLComment SetVisible()
        {
            Visible = true;
            return Container;
        }

        public IXLComment SetVisible(bool hidden)
        {
            Visible = hidden;
            return Container;
        }

        public IXLDrawingPosition Position { get; private set; }

        public int ZOrder { get; set; }

        public IXLComment SetZOrder(int zOrder)
        {
            ZOrder = zOrder;
            return Container;
        }

        public IXLDrawingStyle Style { get; private set; }

        public IXLComment SetName(string name)
        {
            Name = name;
            return Container;
        }

        public IXLComment SetDescription(string description)
        {
            Description = description;
            return Container;
        }

        public IXLComment SetHorizontalFlip()
        {
            HorizontalFlip = true;
            return Container;
        }

        public IXLComment SetHorizontalFlip(bool horizontalFlip)
        {
            HorizontalFlip = horizontalFlip;
            return Container;
        }

        public IXLComment SetVerticalFlip()
        {
            VerticalFlip = true;
            return Container;
        }

        public IXLComment SetVerticalFlip(bool verticalFlip)
        {
            VerticalFlip = verticalFlip;
            return Container;
        }

        public IXLComment SetRotation(int rotation)
        {
            Rotation = rotation;
            return Container;
        }

        public IXLComment SetExtentLength(int extentLength)
        {
            ExtentLength = extentLength;
            return Container;
        }

        public IXLComment SetExtentWidth(int extentWidth)
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
            var previousRowNumber = cell.Address.RowNumber;
            double previousRowOffset = 0;

            if (previousRowNumber > 1)
            {
                previousRowNumber--;

                if (cell.Worksheet.Internals.RowsCollection.TryGetValue(previousRowNumber, out var previousRow))
                {
                    previousRowOffset = Math.Max(0, previousRow.Height - 7);
                }
                else
                {
                    previousRowOffset = Math.Max(0, cell.Worksheet.RowHeight - 7);
                }
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