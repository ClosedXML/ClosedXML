using ClosedXML.Excel.Caching;

namespace ClosedXML.Excel
{
    internal sealed class XLFillValue
    {
        private static readonly XLFillRepository Repository = new XLFillRepository(key => new XLFillValue(key));

        public static XLFillValue FromKey(XLFillKey key)
        {
            return Repository.GetOrCreate(key);
        }

        internal static readonly XLFillValue Default = FromKey(new XLFillKey
        {
            BackgroundColor = XLColor.FromIndex(64).Key,
            PatternType = XLFillPatternValues.None,
            PatternColor = XLColor.FromIndex(64).Key
        });

        public XLFillKey Key { get; private set; }

        public XLColor BackgroundColor { get; private set; }

        public XLColor PatternColor { get; private set; }

        public XLFillPatternValues PatternType { get { return Key.PatternType; } }

        private XLFillValue(XLFillKey key)
        {
            Key = key;
            BackgroundColor = XLColor.FromKey(Key.BackgroundColor);
            PatternColor = XLColor.FromKey(Key.PatternColor);
        }

        public override bool Equals(object obj)
        {
            var cached = obj as XLFillValue;
            return cached != null &&
                   Key.Equals(cached.Key);
        }

        public override int GetHashCode()
        {
            return -280332839 + Key.GetHashCode();
        }
    }
}
