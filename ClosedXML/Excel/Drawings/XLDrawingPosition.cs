namespace ClosedXML.Excel
{
    internal class XLDrawingPosition: IXLDrawingPosition
    {
        public int Column { get; set; }
        public IXLDrawingPosition SetColumn(int column) { Column = column; return this; }
        public double ColumnOffset { get; set; }
        public IXLDrawingPosition SetColumnOffset(double columnOffset) { ColumnOffset = columnOffset; return this; }
         
        public int Row { get; set; }
        public IXLDrawingPosition SetRow(int row) { Row = row; return this; }
        public double RowOffset { get; set; }
        public IXLDrawingPosition SetRowOffset(double rowOffset) { RowOffset = rowOffset; return this; }
    }
}
