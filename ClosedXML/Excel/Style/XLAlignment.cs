using System;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLAlignment : IXLAlignment
    {
        #region Static members

        internal static XLAlignmentKey GenerateKey(IXLAlignment? d)
        {
            XLAlignmentKey key;
            if (d == null)
            {
                key = XLAlignmentValue.Default.Key;
            }
            else if (d is XLAlignment alignment)
            {
                key = alignment.Key;
            }
            else
            {
                key = new XLAlignmentKey
                {
                    Horizontal = d.Horizontal,
                    Vertical = d.Vertical,
                    Indent = d.Indent,
                    JustifyLastLine = d.JustifyLastLine,
                    ReadingOrder = d.ReadingOrder,
                    RelativeIndent = d.RelativeIndent,
                    ShrinkToFit = d.ShrinkToFit,
                    TextRotation = d.TextRotation,
                    WrapText = d.WrapText
                };
            }
            return key;
        }

        #endregion Static members

        #region Properties
        private readonly XLStyle _style;

        private XLAlignmentValue _value;

        internal XLAlignmentKey Key
        {
            get { return _value.Key; }
            private set { _value = XLAlignmentValue.FromKey(ref value); }
        }

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Create an instance of XLAlignment initializing it with the specified value.
        /// </summary>
        /// <param name="style">Style to attach the new instance to.</param>
        /// <param name="value">Style value to use.</param>
        public XLAlignment(XLStyle? style, XLAlignmentValue value)
        {
            _style = style ?? XLStyle.CreateEmptyStyle();
            _value = value;
        }

        public XLAlignment(XLStyle? style, XLAlignmentKey key) : this(style, XLAlignmentValue.FromKey(ref key))
        {
        }

        public XLAlignment(XLStyle? style = null, IXLAlignment? d = null) : this(style, GenerateKey(d))
        {
        }

        #endregion Constructors

        #region IXLAlignment Members

        public XLAlignmentHorizontalValues Horizontal
        {
            get { return Key.Horizontal; }
            set
            {
                Boolean updateIndent = !(
                                               value == XLAlignmentHorizontalValues.Left
                                            || value == XLAlignmentHorizontalValues.Right
                                            || value == XLAlignmentHorizontalValues.Distributed
                                        );

                Modify(k => k with { Horizontal = value });
                if (updateIndent)
                    Indent = 0;
            }
        }

        public XLAlignmentVerticalValues Vertical
        {
            get { return Key.Vertical; }
            set { Modify(k => k with { Vertical = value }); }
        }

        public Int32 Indent
        {
            get { return Key.Indent; }
            set
            {
                if (Indent != value)
                {
                    if (Horizontal == XLAlignmentHorizontalValues.General)
                        Horizontal = XLAlignmentHorizontalValues.Left;

                    if (value > 0 && !(
                                       Horizontal == XLAlignmentHorizontalValues.Left
                                    || Horizontal == XLAlignmentHorizontalValues.Right
                                    || Horizontal == XLAlignmentHorizontalValues.Distributed
                                ))
                    {
                        throw new ArgumentException(
                            "For indents, only left, right, and distributed horizontal alignments are supported.");
                    }
                }
                Modify(k => k with { Indent = value });
            }
        }

        public Boolean JustifyLastLine
        {
            get { return Key.JustifyLastLine; }
            set { Modify(k => k with { JustifyLastLine = value }); }
        }

        public XLAlignmentReadingOrderValues ReadingOrder
        {
            get { return Key.ReadingOrder; }
            set { Modify(k => k with { ReadingOrder = value }); }
        }

        public Int32 RelativeIndent
        {
            get { return Key.RelativeIndent; }
            set { Modify(k => k with { RelativeIndent = value }); }
        }

        public Boolean ShrinkToFit
        {
            get { return Key.ShrinkToFit; }
            set { Modify(k => k with { ShrinkToFit = value }); }
        }

        public Int32 TextRotation
        {
            get { return Key.TextRotation; }
            set
            {
                Int32 rotation = value;

                if (rotation != 255 && (rotation < -90 || rotation > 90))
                    throw new ArgumentException("TextRotation must be between -90 and 90 degrees, or 255.");

                Modify(k => k with { TextRotation = rotation });
            }
        }

        public Boolean WrapText
        {
            get { return Key.WrapText; }
            set { Modify(k => k with { WrapText = value }); }
        }

        public Boolean TopToBottom
        {
            get { return TextRotation == 255; }
            set { TextRotation = value ? 255 : 0; }
        }

        public IXLStyle SetHorizontal(XLAlignmentHorizontalValues value)
        {
            Horizontal = value;
            return _style;
        }

        public IXLStyle SetVertical(XLAlignmentVerticalValues value)
        {
            Vertical = value;
            return _style;
        }

        public IXLStyle SetIndent(Int32 value)
        {
            Indent = value;
            return _style;
        }

        public IXLStyle SetJustifyLastLine()
        {
            JustifyLastLine = true;
            return _style;
        }

        public IXLStyle SetJustifyLastLine(Boolean value)
        {
            JustifyLastLine = value;
            return _style;
        }

        public IXLStyle SetReadingOrder(XLAlignmentReadingOrderValues value)
        {
            ReadingOrder = value;
            return _style;
        }

        public IXLStyle SetRelativeIndent(Int32 value)
        {
            RelativeIndent = value;
            return _style;
        }

        public IXLStyle SetShrinkToFit()
        {
            ShrinkToFit = true;
            return _style;
        }

        public IXLStyle SetShrinkToFit(Boolean value)
        {
            ShrinkToFit = value;
            return _style;
        }

        public IXLStyle SetTextRotation(Int32 value)
        {
            TextRotation = value;
            return _style;
        }

        public IXLStyle SetWrapText()
        {
            WrapText = true;
            return _style;
        }

        public IXLStyle SetWrapText(Boolean value)
        {
            WrapText = value;
            return _style;
        }

        public IXLStyle SetTopToBottom()
        {
            TopToBottom = true;
            return _style;
        }

        public IXLStyle SetTopToBottom(Boolean value)
        {
            TopToBottom = value;
            return _style;
        }

        #endregion

        private void Modify(Func<XLAlignmentKey, XLAlignmentKey> modification)
        {
            Key = modification(Key);

            _style.Modify(styleKey =>
            {
                var alignment = modification(styleKey.Alignment);
                return styleKey with { Alignment = alignment };
            });
        }

        #region Overridden

        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.Append(Horizontal);
            sb.Append("-");
            sb.Append(Vertical);
            sb.Append("-");
            sb.Append(Indent);
            sb.Append("-");
            sb.Append(JustifyLastLine);
            sb.Append("-");
            sb.Append(ReadingOrder);
            sb.Append("-");
            sb.Append(RelativeIndent);
            sb.Append("-");
            sb.Append(ShrinkToFit);
            sb.Append("-");
            sb.Append(TextRotation);
            sb.Append("-");
            sb.Append(WrapText);
            sb.Append("-");
            return sb.ToString();
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as XLAlignment);
        }

        public bool Equals(IXLAlignment? other)
        {
            var otherA = other as XLAlignment;
            if (otherA == null)
                return false;

            return Key == otherA.Key;
        }

        public override int GetHashCode()
        {
            var hashCode = 1214962009;
            hashCode = hashCode * -1521134295 + Key.GetHashCode();
            return hashCode;
        }

        #endregion Overridden
    }
}
