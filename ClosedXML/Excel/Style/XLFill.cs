using System;

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
            private set { _value = XLFillValue.FromKey(ref value); }
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

        public XLFill(XLStyle style, XLFillKey key) : this(style, XLFillValue.FromKey(ref key))
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
            get
            {
                var backgroundColorKey = Key.BackgroundColor;
                return XLColor.FromKey(ref backgroundColorKey);
            }
            set
            {
                if (value == null)
                {
                    throw new ArgumentNullException(nameof(value), "Color cannot be null");
                }

                if ((PatternType == XLFillPatternValues.None ||
                     PatternType == XLFillPatternValues.Solid)
                    && XLColor.IsNullOrTransparent(BackgroundColor))
                {
                    var patternType = value.HasValue ? XLFillPatternValues.Solid : XLFillPatternValues.None;
                    Modify(k =>
                    {
                        k.BackgroundColor = value.Key;
                        k.PatternType = patternType;
                        return k;
                    });
                }
                else
                {
                    Modify(k => { k.BackgroundColor = value.Key; return k; });
                }
            }
        }

        public XLColor PatternColor
        {
            get
            {
                var patternColorKey = Key.PatternColor;
                return XLColor.FromKey(ref patternColorKey);
            }
            set
            {
                if (value == null)
                {
                    throw new ArgumentNullException(nameof(value), "Color cannot be null");
                }

                Modify(k => { k.PatternColor = value.Key; return k; });
            }
        }

        public XLFillPatternValues PatternType
        {
            get { return Key.PatternType; }
            set
            {
                if (PatternType == XLFillPatternValues.None &&
                    value != XLFillPatternValues.None)
                {
                    // If fill was empty and the pattern changes to non-empty we have to specify a background color too.
                    // Otherwise the fill will be considered empty and pattern won't update (the cached empty fill will be used).
                    Modify(k =>
                    {
                        k.BackgroundColor = XLColor.FromTheme(XLThemeColor.Text1).Key;
                        k.PatternType = value;
                        return k;
                    });
                }
                else
                {
                    Modify(k => { k.PatternType = value; return k; });
                }
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
            if (!(other is XLFill otherF))
            {
                return false;
            }

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
