using ClosedXML.Excel.Caching;

namespace ClosedXML.Excel
{
    public class XLAlignmentValue
    {
        private static readonly XLAlignmentRepository Repository = new XLAlignmentRepository(key => new XLAlignmentValue(key));

        public static XLAlignmentValue FromKey(ref XLAlignmentKey key)
        {
            return Repository.GetOrCreate(ref key);
        }

        private static readonly XLAlignmentKey DefaultKey = new XLAlignmentKey
        {
            Indent = 0,
            Horizontal = XLAlignmentHorizontalValues.General,
            JustifyLastLine = false,
            ReadingOrder = XLAlignmentReadingOrderValues.ContextDependent,
            RelativeIndent = 0,
            ShrinkToFit = false,
            TextRotation = 0,
            Vertical = XLAlignmentVerticalValues.Bottom,
            WrapText = false
        };

        internal static readonly XLAlignmentValue Default = FromKey(ref DefaultKey);

        public XLAlignmentKey Key { get; private set; }

        public XLAlignmentHorizontalValues Horizontal { get { return Key.Horizontal; } }

        public XLAlignmentVerticalValues Vertical { get { return Key.Vertical; } }

        public int Indent { get { return Key.Indent; } }

        public bool JustifyLastLine { get { return Key.JustifyLastLine; } }

        public XLAlignmentReadingOrderValues ReadingOrder { get { return Key.ReadingOrder; } }

        public int RelativeIndent { get { return Key.RelativeIndent; } }

        public bool ShrinkToFit { get { return Key.ShrinkToFit; } }

        public int TextRotation { get { return Key.TextRotation; } }

        public bool WrapText { get { return Key.WrapText; } }

        public bool TopToBottom { get { return Key.TopToBottom; } }

        private XLAlignmentValue(XLAlignmentKey key)
        {
            Key = key;
        }

        public override bool Equals(object obj)
        {
            var cached = obj as XLAlignmentValue;
            return cached != null &&
                   Key.Equals(cached.Key);
        }

        public override int GetHashCode()
        {
            return 990326508 + Key.GetHashCode();
        }
    }
}
