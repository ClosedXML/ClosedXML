using System;

namespace ClosedXML.Excel
{
    public enum XLDrawingAnchor { MoveAndSizeWithCells, MoveWithCells, Absolute}
    public interface IXLDrawing<T>
    {
        Int32 ShapeId { get; }

        Boolean Visible { get; set; }
        T SetVisible();
        T SetVisible(Boolean hidden);
                
        String Name { get; set; }
        T SetName(String name);

        String Description { get; set; }
        T SetDescription(String description);
        
        

        Int32 FirstColumn { get; set; }
        T SetFirstColumn(Int32 firstColumn);
        Int32 FirstColumnOffset { get; set; }
        T SetFirstColumnOffset(Int32 firstColumnOffset);
        Int32 FirstRow { get; set; }
        T SetFirstRow(Int32 firstRow);
        Int32 FirstRowOffset { get; set; }
        T SetFirstRowOffset(Int32 firstRowOffset);

        Int32 LastColumn { get; set; }
        T SetLastColumn(Int32 firstColumn);
        Int32 LastColumnOffset { get; set; }
        T SetLastColumnOffset(Int32 firstColumn);
        Int32 LastRow { get; set; }
        T SetLastRow(Int32 firstRow);
        Int32 LastRowOffset { get; set; }
        T SetLastRowOffset(Int32 firstRowOffset);

        Int32 ZOrder { get; set; }
        T SetZOrder(Int32 zOrder);

        Boolean HorizontalFlip { get; set; }
        T SetHorizontalFlip();
        T SetHorizontalFlip(Boolean horizontalFlip);

        Boolean VerticalFlip { get; set; }
        T SetVerticalFlip();
        T SetVerticalFlip(Boolean verticalFlip);

        Int32 Rotation { get; set; }
        T SetRotation(Int32 rotation);

        Int32 OffsetX { get; set; }
        T SetOffsetX(Int32 offsetX);

        Int32 OffsetY { get; set; }
        T SetOffsetY(Int32 offsetY);

        Int32 ExtentLength { get; set; }
        T SetExtentLength(Int32 ExtentLength);

        Int32 ExtentWidth { get; set; }
        T SetExtentWidth(Int32 extentWidth);

        IXLDrawingStyle Style { get; }

        //"position:absolute; margin-left:59.25pt;margin-top:1.5pt;width:96pt;height:60pt;z-index:1; visibility:visible";

    }
}
