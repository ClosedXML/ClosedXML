using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLComment : XLFormattedText<IXLComment>, IXLComment
    {
        XLWorksheet _worksheet;
        XLCell _cell;
        public XLComment(XLCell cell, IXLFontBase defaultFont)
            : base(defaultFont)
        {
            Initialize(cell);
        }

        public XLComment(XLCell cell, XLFormattedText<IXLComment> defaultComment, IXLFontBase defaultFont)
            : base(defaultComment, defaultFont) 
        {
            Initialize(cell);
        }

        public XLComment(XLCell cell, String text, IXLFontBase defaultFont)
            : base(text, defaultFont)
        {
            Initialize(cell);
        }

        private void Initialize(XLCell cell)
        {
            Author = Environment.UserName;
            Container = this;
            Anchor = XLDrawingAnchor.MoveAndSizeWithCells;
            Style = new XLDrawingStyle();
            Int32 pRow = cell.Address.RowNumber;
            Double pRowOffset = 0;
            if (pRow > 1)
            {
                pRow--;
                var prevHeight = cell.CellAbove().WorksheetRow().Height;
                if (prevHeight > 7)
                    pRowOffset = prevHeight - 7;
            }
            Position = new XLDrawingPosition
            {
                Column = cell.Address.ColumnNumber + 1,
                ColumnOffset = 2,
                Row = pRow,
                RowOffset = pRowOffset
            };

            ZOrder = cell.Worksheet.ZOrder++;
            Style
                 .Margins.SetLeft(0.1)
                 .Margins.SetRight(0.1)
                 .Margins.SetTop(0.05)
                 .Margins.SetBottom(0.05)
                 .Margins.SetAutomatic()
                 
                 .Size.SetHeight(59.25)
                 .Size.SetWidth(19.2)

                 .ColorsAndLines.SetLineColor(XLColor.Black)
                 .ColorsAndLines.SetFillColor(XLColor.FromArgb(255,255,225))
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

            _cell = cell;
            _worksheet = cell.Worksheet;
            ShapeId = cell.Worksheet.Workbook.ShapeIdManager.GetNext();
        }

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

        public IXLRichString AddNewLine() 
        {
            return AddText(Environment.NewLine);
        }
        
        #region IXLDrawing

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

        public String Name { get; set; }
        public IXLComment SetName(String name)
        {
            Name = name;
            return Container;
        }

        public String Description { get; set; }
        public IXLComment SetDescription(String description)
        {
            Description = description;
            return Container;
        }

        public XLDrawingAnchor Anchor { get; set; }

        public IXLDrawingPosition Position { get; private set; }

        public Int32 ZOrder { get; set; }
        public IXLComment SetZOrder(Int32 zOrder)
        {
            ZOrder = zOrder;
            return Container;
        }

        public Boolean HorizontalFlip { get; set; }
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

        public Boolean VerticalFlip { get; set; }
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

        public Int32 Rotation { get; set; }
        public IXLComment SetRotation(Int32 rotation)
        {
            Rotation = rotation;
            return Container;
        }

        public Int32 ExtentLength { get; set; }
        public IXLComment SetExtentLength(Int32 extentLength)
        {
            ExtentLength = extentLength;
            return Container;
        }

        public Int32 ExtentWidth { get; set; }
        public IXLComment SetExtentWidth(Int32 extentWidth)
        {
            ExtentWidth = extentWidth;
            return Container;
        }

        public IXLDrawingStyle Style { get; private set; }

        #endregion

        public void Delete() { _cell.DeleteComment(); }

    }

}
