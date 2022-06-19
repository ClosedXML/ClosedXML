namespace ClosedXML.Excel
{
    public interface IXLDrawingSize
    {
        bool AutomaticSize { get; set; }
        double Height { get; set; }
        double Width { get; set; }

        IXLDrawingStyle SetAutomaticSize(); IXLDrawingStyle SetAutomaticSize(bool value);
        IXLDrawingStyle SetHeight(double value);
        IXLDrawingStyle SetWidth(double value);

    }
}
