#nullable disable

namespace ClosedXML.Excel
{
    public partial class XLColor
    {
        internal XLColorKey Key { get; private set; }

        private XLColor() : this(new XLColorKey())
        {
            HasValue = false;
        }

        internal XLColor(XLColorKey key)
        {
            Key = key;
            HasValue = true;
        }
    }
}
