using System;

namespace ClosedXML.Excel
{
    internal class XLNumberFormat : IXLNumberFormat
    {
        #region Static members

        internal static XLNumberFormatKey GenerateKey(IXLNumberFormat defaultNumberFormat)
        {
            if (defaultNumberFormat == null)
                return XLNumberFormatValue.Default.Key;

            if (defaultNumberFormat is XLNumberFormat)
                return (defaultNumberFormat as XLNumberFormat).Key;

            return new XLNumberFormatKey
            {
                NumberFormatId = defaultNumberFormat.NumberFormatId,
                Format = defaultNumberFormat.Format
            };
        }

        #endregion Static members

        #region Properties

        private readonly XLStyle _style;

        private XLNumberFormatValue _value;

        internal XLNumberFormatKey Key
        {
            get { return _value.Key; }
            private set { _value = XLNumberFormatValue.FromKey(ref value); }
        }

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Create an instance of XLNumberFormat initializing it with the specified value.
        /// </summary>
        /// <param name="style">Style to attach the new instance to.</param>
        /// <param name="value">Style value to use.</param>
        public XLNumberFormat(XLStyle style, XLNumberFormatValue value)
        {
            _style = style ?? XLStyle.CreateEmptyStyle();
            _value = value;
        }

        public XLNumberFormat(XLStyle style, XLNumberFormatKey key) : this(style, XLNumberFormatValue.FromKey(ref key))
        {
        }

        public XLNumberFormat(XLStyle style = null, IXLNumberFormat d = null) : this(style, GenerateKey(d))
        {
        }

        #endregion Constructors

        #region IXLNumberFormat Members

        public Int32 NumberFormatId
        {
            get { return Key.NumberFormatId; }
            set
            {
                Modify(k =>
                {
                    k.Format = XLNumberFormatValue.Default.Format;
                    k.NumberFormatId = value;
                    return k;
                });
            }
        }

        public String Format
        {
            get { return Key.Format; }
            set
            {
                Modify(k =>
                {
                    k.Format = value;
                    if (string.IsNullOrWhiteSpace(k.Format))
                        k.NumberFormatId = XLNumberFormatValue.Default.NumberFormatId;
                    else
                        k.NumberFormatId = -1;
                    return k;
                });
            }
        }

        public IXLStyle SetNumberFormatId(Int32 value)
        {
            NumberFormatId = value;
            return _style;
        }

        public IXLStyle SetFormat(String value)
        {
            Format = value;
            return _style;
        }

        #endregion IXLNumberFormat Members

        private void Modify(Func<XLNumberFormatKey, XLNumberFormatKey> modification)
        {
            Key = modification(Key);

            _style.Modify(styleKey =>
            {
                var numberFormat = styleKey.NumberFormat;
                styleKey.NumberFormat = modification(numberFormat);
                return styleKey;
            });
        }

        #region Overridden

        public override string ToString()
        {
            return NumberFormatId + "-" + Format;
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as IXLNumberFormatBase);
        }

        public bool Equals(IXLNumberFormatBase other)
        {
            var otherN = other as XLNumberFormat;
            if (otherN == null)
                return false;

            return Key == otherN.Key;
        }

        public override int GetHashCode()
        {
            var hashCode = 416600561;
            hashCode = hashCode * -1521134295 + Key.GetHashCode();
            return hashCode;
        }

        #endregion Overridden
    }
}
