using System;

namespace ClosedXML.Excel
{
    internal class XLDrawing<T>: IXLDrawing<T>
    {
        internal T Container;
        public XLDrawing()
        {
            Style = new XLDrawingStyle();
            Position = new XLDrawingPosition();
        }

        public Int32 ShapeId { get; internal set; }

        public Boolean Visible { get; set; }
        public T SetVisible()
        {
            Visible = true;
            return Container;
        }
        public T SetVisible(Boolean hidden)
        {
            Visible = hidden;
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

        public IXLDrawingPosition Position { get; private set; }

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
