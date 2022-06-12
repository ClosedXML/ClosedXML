using ClosedXML.Excel.Caching;

namespace ClosedXML.Excel
{
    internal sealed class XLNumberFormatValue
    {
        private static readonly XLNumberFormatRepository Repository = new XLNumberFormatRepository(key => new XLNumberFormatValue(key));

        public static XLNumberFormatValue FromKey(ref XLNumberFormatKey key)
        {
            return Repository.GetOrCreate(ref key);
        }

        private static readonly XLNumberFormatKey DefaultKey = new XLNumberFormatKey
        {
            NumberFormatId = 0,
            Format = string.Empty
        };

        internal static readonly XLNumberFormatValue Default = FromKey(ref DefaultKey);

        public XLNumberFormatKey Key { get; private set; }

        public int NumberFormatId => Key.NumberFormatId;

        public string Format => Key.Format;

        private XLNumberFormatValue(XLNumberFormatKey key)
        {
            Key = key;
        }

        public override bool Equals(object obj)
        {
            return obj is XLNumberFormatValue cached &&
                   Key.Equals(cached.Key);
        }

        public override int GetHashCode()
        {
            return 1507230172 + Key.GetHashCode();
        }
    }
}
