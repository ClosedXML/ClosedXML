using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLComment : XLFormattedText<IXLComment>, IXLComment
    {

        public XLComment(IXLFontBase defaultFont)
            : base(defaultFont)
        {
            Initialize();
        }

        public XLComment(XLFormattedText<IXLComment> defaultComment, IXLFontBase defaultFont)
            : base(defaultComment, defaultFont) 
        {
            Initialize();
        }

        public XLComment(String text, IXLFontBase defaultFont)
            : base(text, defaultFont)
        {
            Initialize();
        }

        private void Initialize()
        {
            Container = this;
            Anchor = XLDrawingAnchor.MoveAndSizeWithCells;
            Style = new XLDrawingStyle();
            Style.Size.Height = 4;  // I think this is misused for legacy drawing
            Style.Size.Width = 2;
            SetVisible();
        }

        public String Author { get; set; }
        public IXLComment SetAuthor(String value)
        {
            Author = value;
            return this;
        }

        public IXLRichString AddSignature() 
        {
            // existing Author might be someone else hence using current user name here
            return AddSignature(Environment.UserName);

        }

        public IXLRichString AddSignature(string username) 
        {
            return AddText(string.Format("{0}:{1}", username, Environment.NewLine)).SetBold();
        }

        public IXLRichString AddNewLine() 
        {
            return AddText(Environment.NewLine);
        }

        public Boolean Visible { get; set; }	public IXLComment SetVisible() { Visible = true; return this; }	public IXLComment SetVisible(Boolean value) { Visible = value; return this; }

        #region IXLDrawing

        public Int32 Id { get; internal set; }

        public Boolean Hidden { get; set; }
        public IXLComment SetHidden()
        {
            Hidden = true;
            return Container;
        }
        public IXLComment SetHidden(Boolean hidden)
        {
            Hidden = hidden;
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

        public Int32 FirstColumn { get; set; }
        public IXLComment SetFirstColumn(Int32 firstColumn)
        {
            FirstColumn = firstColumn;
            return Container;
        }
        public Int32 FirstColumnOffset { get; set; }
        public IXLComment SetFirstColumnOffset(Int32 firstColumnOffset)
        {
            FirstColumnOffset = firstColumnOffset;
            return Container;
        }
        public Int32 FirstRow { get; set; }
        public IXLComment SetFirstRow(Int32 firstRow)
        {
            FirstRow = firstRow;
            return Container;
        }
        public Int32 FirstRowOffset { get; set; }
        public IXLComment SetFirstRowOffset(Int32 firstRowOffset)
        {
            FirstRowOffset = firstRowOffset;
            return Container;
        }

        public Int32 LastColumn { get; set; }
        public IXLComment SetLastColumn(Int32 firstColumn)
        {
            LastColumn = firstColumn;
            return Container;
        }
        public Int32 LastColumnOffset { get; set; }
        public IXLComment SetLastColumnOffset(Int32 firstColumnOffset)
        {
            LastColumnOffset = firstColumnOffset;
            return Container;
        }
        public Int32 LastRow { get; set; }
        public IXLComment SetLastRow(Int32 firstRow)
        {
            LastRow = firstRow;
            return Container;
        }
        public Int32 LastRowOffset { get; set; }
        public IXLComment SetLastRowOffset(Int32 firstRowOffset)
        {
            LastRowOffset = firstRowOffset;
            return Container;
        }

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

        public Int32 OffsetX { get; set; }
        public IXLComment SetOffsetX(Int32 offsetX)
        {
            OffsetX = offsetX;
            return Container;
        }

        public Int32 OffsetY { get; set; }
        public IXLComment SetOffsetY(Int32 offsetY)
        {
            OffsetY = offsetY;
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

    }

}
