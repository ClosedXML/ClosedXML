using System;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLFill : IXLFill
    {
        #region static members

        internal static XLFillKey GenerateKey(IXLFill defaultFill)
        {
            XLFillKey key;
            if (defaultFill == null)
            {
                key = XLFillValue.Default.Key;
            }
            else if (defaultFill is XLFill)
            {
                key = (defaultFill as XLFill).Key;
            }
            else
            {
                key = new XLFillKey
                {
                    PatternType = defaultFill.PatternType,
                    BackgroundColor = defaultFill.BackgroundColor.Key,
                    PatternColor = defaultFill.PatternColor.Key
                };
            }
            return key;
        }

        #endregion static members

        #region Properties

        private readonly XLStyle _style;

        private XLFillValue _value;

        internal XLFillKey Key
        {
            get { return _value.Key; }
            private set { _value = XLFillValue.FromKey(value); }
        }

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Create an instance of XLFill initializing it with the specified value.
        /// </summary>
        /// <param name="style">Style to attach the new instance to.</param>
        /// <param name="value">Style value to use.</param>
        public XLFill(XLStyle style, XLFillValue value)
        {
            _style = style ?? XLStyle.CreateEmptyStyle();
            _value = value;
        }

        public XLFill(XLStyle style, XLFillKey key) : this(style, XLFillValue.FromKey(key))
        {
        }

        public XLFill(XLStyle style = null, IXLFill d = null) : this(style, GenerateKey(d))
        {
        }

        #endregion Constructors

        private void Modify(Func<XLFillKey, XLFillKey> modification)
        {
            Key = modification(Key);

            _style.Modify(styleKey =>
            {
                var fill = styleKey.Fill;
                styleKey.Fill = modification(fill);
                return styleKey;
            });
        }

        #region IXLFill Members

        public XLColor BackgroundColor
        {
            get { return XLColor.FromKey(Key.BackgroundColor); }
            set
            {
                if (value == null)
                    throw new ArgumentNullException(nameof(value), "Color cannot be null");

                // 4 ways of determining an "empty" color
                if (new XLFillPatternValues[] { XLFillPatternValues.None, XLFillPatternValues.Solid }.Contains(PatternType)
                    && (BackgroundColor == null
                    || !BackgroundColor.HasValue
                    || BackgroundColor == XLColor.NoColor
                    || BackgroundColor.ColorType == XLColorType.Indexed && BackgroundColor.Indexed == 64))
                {
                    PatternType = value.HasValue ? XLFillPatternValues.Solid : XLFillPatternValues.None;
                }

                Modify(k => { k.BackgroundColor = value.Key; return k; });
            }
        }

        public XLColor PatternColor
        {
            get { return XLColor.FromKey(Key.PatternColor); }
            set
            {
                if (value == null)
                    throw new ArgumentNullException(nameof(value), "Color cannot be null");

                Modify(k => { k.PatternColor = value.Key; return k; });
            }
        }

        public XLFillPatternValues PatternType
        {
            get { return Key.PatternType; }
            set
            {
                Modify(k => { k.PatternType = value; return k; });
            }
        }

        public IXLStyle SetBackgroundColor(XLColor value)
        {
            BackgroundColor = value;
            return _style;
        }

        public IXLStyle SetPatternColor(XLColor value)
        {
            PatternColor = value;
            return _style;
        }

        public IXLStyle SetPatternType(XLFillPatternValues value)
        {
            PatternType = value;
            return _style;
        }

        #endregion IXLFill Members

        #region Overridden

        public override bool Equals(object obj)
        {
            return Equals(obj as XLFill);
        }

        public bool Equals(IXLFill other)
        {
            var otherF = other as XLFill;
            if (otherF == null)
                return false;

            return Key == otherF.Key;
        }

        public override string ToString()
        {
            switch (PatternType)
            {
                case XLFillPatternValues.None:
                    return "None";

                case XLFillPatternValues.Solid:
                    return string.Concat("Solid ", BackgroundColor.ToString());

                default:
                    return string.Concat(PatternType.ToString(), " pattern: ", PatternColor.ToString(), " on ", BackgroundColor.ToString());
            }
        }

        public override int GetHashCode()
        {
            var hashCode = -1938644919;
            hashCode = hashCode * -1521134295 + Key.GetHashCode();
            return hashCode;
        }

        #endregion Overridden
    }
}
