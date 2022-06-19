namespace ClosedXML.Excel
{
    public interface IXLDrawingProperties
    {
        XLDrawingAnchor Positioning { get; set; }
        IXLDrawingStyle SetPositioning(XLDrawingAnchor value);

    }
}
