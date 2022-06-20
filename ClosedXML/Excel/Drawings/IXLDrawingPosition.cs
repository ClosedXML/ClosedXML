namespace ClosedXML.Excel
{
    public interface IXLDrawingPosition
    {
        int Column { get; set; }
        IXLDrawingPosition SetColumn(int column);
        double ColumnOffset { get; set; }
        IXLDrawingPosition SetColumnOffset(double columnOffset);

        int Row { get; set; }
        IXLDrawingPosition SetRow(int row);
        double RowOffset { get; set; }
        IXLDrawingPosition SetRowOffset(double rowOffset);
    }
}
