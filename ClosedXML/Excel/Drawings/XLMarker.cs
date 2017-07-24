using System.Diagnostics;
using System.Drawing;

namespace ClosedXML.Excel.Drawings
{
    [DebuggerDisplay("{Address} {Offset}")]
    internal class XLMarker : IXLMarker
    {
        internal XLMarker(IXLAddress address)
            : this(address, new Point(0, 0))
        { }

        internal XLMarker(IXLAddress address, Point offset)
        {
            this.Address = address;
            this.Offset = offset;
        }

        public IXLAddress Address { get; set; }

        public Point Offset { get; set; }
    }
}
