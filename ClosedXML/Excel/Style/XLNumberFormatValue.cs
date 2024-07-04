#nullable disable

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

        /// <summary>
        /// <em>General</em> number format.
        /// </summary>
        private static readonly XLNumberFormatKey DefaultKey = new XLNumberFormatKey
        {
            NumberFormatId = 0,
            Format = string.Empty
        };

        internal static readonly XLNumberFormatValue Default = FromKey(ref DefaultKey);

        public XLNumberFormatKey Key { get; private set; }

        /// <summary>
        /// Id of the number format. Every workbook has <see cref="XLConstants.NumberOfBuiltInStyles"/>
        /// built-int formats that start at 0 (<em>General</em> format). The built-int formats are
        /// not explicitly written and might differ depending on culture. Custom number formats
        /// have a valid <see cref="Format"/> and the id is <c>-1</c>.
        /// </summary>
        public int NumberFormatId { get { return Key.NumberFormatId; } }

        public string Format { get { return Key.Format; } }

        private XLNumberFormatValue(XLNumberFormatKey key)
        {
            Key = key;
        }

        public override bool Equals(object obj)
        {
            var cached = obj as XLNumberFormatValue;
            return cached != null &&
                   Key.Equals(cached.Key);
        }

        public override int GetHashCode()
        {
            return 1507230172 + Key.GetHashCode();
        }

        internal XLNumberFormatValue WithNumberFormatId(int numberFormatId)
        {
            var keyCopy = Key with { NumberFormatId = numberFormatId };
            return FromKey(ref keyCopy);
        }
    }
}
