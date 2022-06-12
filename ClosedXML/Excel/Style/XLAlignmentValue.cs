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

        public XLAlignmentHorizontalValues Horizontal => Key.Horizontal;

        public XLAlignmentVerticalValues Vertical => Key.Vertical;

        public int Indent => Key.Indent;

        public bool JustifyLastLine => Key.JustifyLastLine;

        public XLAlignmentReadingOrderValues ReadingOrder => Key.ReadingOrder;

        public int RelativeIndent => Key.RelativeIndent;

        public bool ShrinkToFit => Key.ShrinkToFit;

        public int TextRotation => Key.TextRotation;

        public bool WrapText => Key.WrapText;

        public bool TopToBottom => Key.TopToBottom;

        private XLAlignmentValue(XLAlignmentKey key)
        {
            Key = key;
        }

        public override bool Equals(object obj)
        {
            return obj is XLAlignmentValue cached &&
                   Key.Equals(cached.Key);
        }

        public override int GetHashCode()
        {
            return 990326508 + Key.GetHashCode();
        }
    }
}
