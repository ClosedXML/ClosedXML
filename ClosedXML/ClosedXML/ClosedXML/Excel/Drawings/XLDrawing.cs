using System;

namespace ClosedXML.Excel
{
    internal class XLDrawing<T>: IXLDrawing<T>
    {
        internal T Container;
        public XLDrawing()
        {
            Anchor = XLDrawingAnchor.MoveAndSizeWithCells;
            Style = new XLDrawingStyle();
        }

        public Int32 Id { get; internal set; }

        public Boolean Hidden { get; set; }
        public T SetHidden()
        {
            Hidden = true;
            return Container;
        }
        public T SetHidden(Boolean hidden)
        {
            Hidden = hidden;
            return Container;
        }

        public String Name { get; set; }
        public T SetName(String name)
        {
            Name = name;
            return Container;
        }

        public String Description { get; set; }
        public T SetDescription(String description)
        {
            Description = description;
            return Container;
        }

        public XLDrawingAnchor Anchor { get; set; }

        public Int32 FirstColumn { get; set; }
        public T SetFirstColumn(Int32 firstColumn)
        {
            FirstColumn = firstColumn;
            return Container;
        }
        public Int32 FirstColumnOffset { get; set; }
        public T SetFirstColumnOffset(Int32 firstColumnOffset)
        {
            FirstColumnOffset = firstColumnOffset;
            return Container;
        }
        public Int32 FirstRow { get; set; }
        public T SetFirstRow(Int32 firstRow)
        {
            FirstRow = firstRow;
            return Container;
        }
        public Int32 FirstRowOffset { get; set; }
        public T SetFirstRowOffset(Int32 firstRowOffset)
        {
            FirstRowOffset = firstRowOffset;
            return Container;
        }

        public Int32 LastColumn { get; set; }
        public T SetLastColumn(Int32 firstColumn)
        {
            LastColumn = firstColumn;
            return Container;
        }
        public Int32 LastColumnOffset { get; set; }
        public T SetLastColumnOffset(Int32 firstColumnOffset)
        {
            LastColumnOffset = firstColumnOffset;
            return Container;
        }
        public Int32 LastRow { get; set; }
        public T SetLastRow(Int32 firstRow)
        {
            LastRow = firstRow;
            return Container;
        }
        public Int32 LastRowOffset { get; set; }
        public T SetLastRowOffset(Int32 firstRowOffset)
        {
            LastRowOffset = firstRowOffset;
            return Container;
        }

        public Int32 ZOrder { get; set; }
        public T SetZOrder(Int32 zOrder)
        {
            ZOrder = zOrder;
            return Container;
        }

        public Boolean HorizontalFlip { get; set; }
        public T SetHorizontalFlip()
        {
            HorizontalFlip = true;
            return Container;
        }
        public T SetHorizontalFlip(Boolean horizontalFlip)
        {
            HorizontalFlip = horizontalFlip;
            return Container;
        }

        public Boolean VerticalFlip { get; set; }
        public T SetVerticalFlip()
        {
            VerticalFlip = true;
            return Container;
        }
        public T SetVerticalFlip(Boolean verticalFlip)
        {
            VerticalFlip = verticalFlip;
            return Container;
        }

        public Int32 Rotation { get; set; }
        public T SetRotation(Int32 rotation)
        {
            Rotation = rotation;
            return Container;
        }

        public Int32 OffsetX { get; set; }
        public T SetOffsetX(Int32 offsetX)
        {
            OffsetX = offsetX;
            return Container;
        }

        public Int32 OffsetY { get; set; }
        public T SetOffsetY(Int32 offsetY)
        {
            OffsetY = offsetY;
            return Container;
        }

        public Int32 ExtentLength { get; set; }
        public T SetExtentLength(Int32 extentLength)
        {
            ExtentLength = extentLength;
            return Container;
        }

        public Int32 ExtentWidth { get; set; }
        public T SetExtentWidth(Int32 extentWidth)
        {
            ExtentWidth = extentWidth;
            return Container;
        }

        public IXLDrawingStyle Style { get; private set; }
    }
}
