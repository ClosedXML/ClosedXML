using ClosedXML.Excel.Caching;

namespace ClosedXML.Excel
{
    internal class XLBorderValue
    {
        private static readonly XLBorderRepository Repository = new XLBorderRepository(key => new XLBorderValue(key));

        public static XLBorderValue FromKey(XLBorderKey key)
        {
            return Repository.GetOrCreate(key);
        }

        internal static readonly XLBorderValue Default = FromKey(new XLBorderKey
        {
            BottomBorder = XLBorderStyleValues.None,
            DiagonalBorder = XLBorderStyleValues.None,
            DiagonalDown = false,
            DiagonalUp = false,
            LeftBorder = XLBorderStyleValues.None,
            RightBorder = XLBorderStyleValues.None,
            TopBorder = XLBorderStyleValues.None,
            BottomBorderColor = XLColor.Black.Key,
            DiagonalBorderColor = XLColor.Black.Key,
            LeftBorderColor = XLColor.Black.Key,
            RightBorderColor = XLColor.Black.Key,
            TopBorderColor = XLColor.Black.Key
        });

        public XLBorderKey Key { get; private set; }

        public XLBorderStyleValues LeftBorder { get { return Key.LeftBorder; } }

        public XLColor LeftBorderColor { get; private set; }

        public XLBorderStyleValues RightBorder { get { return Key.RightBorder; } }

        public XLColor RightBorderColor { get; private set; }

        public XLBorderStyleValues TopBorder { get { return Key.TopBorder; } }

        public XLColor TopBorderColor { get; private set; }

        public XLBorderStyleValues BottomBorder { get { return Key.BottomBorder; } }

        public XLColor BottomBorderColor { get; private set; }

        public XLBorderStyleValues DiagonalBorder { get { return Key.DiagonalBorder; } }

        public XLColor DiagonalBorderColor { get; private set; }

        public bool DiagonalUp { get { return Key.DiagonalUp; } }

        public bool DiagonalDown { get { return Key.DiagonalDown; } }

        private XLBorderValue(XLBorderKey key)
        {
            Key = key;

            LeftBorderColor = XLColor.FromKey(Key.LeftBorderColor);
            RightBorderColor = XLColor.FromKey(Key.RightBorderColor);
            TopBorderColor = XLColor.FromKey(Key.TopBorderColor);
            BottomBorderColor = XLColor.FromKey(Key.BottomBorderColor);
            DiagonalBorderColor = XLColor.FromKey(Key.DiagonalBorderColor);
        }

        public override bool Equals(object obj)
        {
            var cached = obj as XLBorderValue;
            return cached != null &&
                   Key.Equals(cached.Key);
        }

        public override int GetHashCode()
        {
            return -280332839 + Key.GetHashCode();
        }
    }
}
