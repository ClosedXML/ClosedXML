namespace ClosedXML.Excel
{
    public enum XLDrawingAnchor { MoveAndSizeWithCells, MoveWithCells, Absolute}
    public interface IXLDrawing<T>
    {
        int ShapeId { get; }

        bool Visible { get; set; }
        T SetVisible();
        T SetVisible(bool hidden);
                
        ////String Name { get; set; }
        ////T SetName(String name);

        ////String Description { get; set; }
        ////T SetDescription(String description);

        IXLDrawingPosition Position { get;  }

        int ZOrder { get; set; }
        T SetZOrder(int zOrder);

        //Boolean HorizontalFlip { get; set; }
        //T SetHorizontalFlip();
        //T SetHorizontalFlip(Boolean horizontalFlip);

        //Boolean VerticalFlip { get; set; }
        //T SetVerticalFlip();
        //T SetVerticalFlip(Boolean verticalFlip);

        //Int32 Rotation { get; set; }
        //T SetRotation(Int32 rotation);

        //Int32 ExtentLength { get; set; }
        //T SetExtentLength(Int32 ExtentLength);

        //Int32 ExtentWidth { get; set; }
        //T SetExtentWidth(Int32 extentWidth);

        IXLDrawingStyle Style { get; }

    }
}
