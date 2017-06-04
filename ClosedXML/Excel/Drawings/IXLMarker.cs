using System.Drawing;

namespace ClosedXML.Excel.Drawings
{
    internal interface IXLMarker
    {
        IXLAddress Address { get; set; }
        Point Offset { get; set; }
    }
}
