using ClosedXML.Excel.Caching;

namespace ClosedXML.Excel
{
    internal sealed class XLProtectionValue
    {
        private static readonly XLProtectionRepository Repository = new XLProtectionRepository(key => new XLProtectionValue(key));

        public static XLProtectionValue FromKey(XLProtectionKey key)
        {
            return Repository.GetOrCreate(key);
        }

        internal static readonly XLProtectionValue Default = FromKey(new XLProtectionKey
        {
            Locked = true,
            Hidden = false
        });

        public XLProtectionKey Key { get; private set; }

        public bool Locked { get { return Key.Locked; } }

        public bool Hidden { get { return Key.Hidden; } }

        private XLProtectionValue(XLProtectionKey key)
        {
            Key = key;
        }

        public override bool Equals(object obj)
        {
            var cached = obj as XLProtectionValue;
            return cached != null &&
                   Key.Equals(cached.Key);
        }

        public override int GetHashCode()
        {
            return 909014992 + Key.GetHashCode();
        }
    }
}
