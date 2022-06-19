namespace ClosedXML.Excel
{
    public interface IXLDrawingMargins
    {
        bool Automatic { get; set; }
        double Left { get; set; }
        double Right { get; set; }
        double Top { get; set; }
        double Bottom { get; set; }
        double All { set; }

        IXLDrawingStyle SetAutomatic(); IXLDrawingStyle SetAutomatic(bool value);
        IXLDrawingStyle SetLeft(double value);
        IXLDrawingStyle SetRight(double value);
        IXLDrawingStyle SetTop(double value);
        IXLDrawingStyle SetBottom(double value);
        IXLDrawingStyle SetAll(double value);

    }
}
