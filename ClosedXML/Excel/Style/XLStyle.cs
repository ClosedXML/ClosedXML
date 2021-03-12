using System;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLStyle : IXLStyle
    {
        #region Static members

        public static XLStyle Default { get { return new XLStyle(XLStyleValue.Default); } }

        internal static XLStyleKey GenerateKey(IXLStyle initialStyle)
        {
            if (initialStyle == null)
                return Default.Key;
            if (initialStyle is XLStyle)
                return (initialStyle as XLStyle).Key;

            return new XLStyleKey
            {
                Font = XLFont.GenerateKey(initialStyle.Font),
                Alignment = XLAlignment.GenerateKey(initialStyle.Alignment),
                Border = XLBorder.GenerateKey(initialStyle.Border),
                Fill = XLFill.GenerateKey(initialStyle.Fill),
                NumberFormat = XLNumberFormat.GenerateKey(initialStyle.NumberFormat),
                Protection = XLProtection.GenerateKey(initialStyle.Protection)
            };
        }

        internal static XLStyle CreateEmptyStyle()
        {
            return new XLStyle(new XLStylizedEmpty(null));
        }

        #endregion Static members

        #region properties

        private readonly IXLStylized _container;

        internal XLStyleValue Value { get; private set; }

        internal XLStyleKey Key
        {
            get { return Value.Key; }
            private set
            {
                Value = XLStyleValue.FromKey(ref value);
            }
        }

        #endregion properties

        #region constructors

        public XLStyle(IXLStylized container, IXLStyle initialStyle = null, Boolean useDefaultModify = true) : this(container, GenerateKey(initialStyle))
        {
        }

        public XLStyle(IXLStylized container, XLStyleKey key) : this(container, XLStyleValue.FromKey(ref key))
        {
        }

        internal XLStyle(IXLStylized container, XLStyleValue value)
        {
            _container = container ?? new XLStylizedEmpty(XLStyle.Default);
            Value = value;
        }

        /// <summary>
        /// To initialize XLStyle.Default only
        /// </summary>
        private XLStyle(XLStyleValue value)
        {
            _container = null;
            Value = value;
        }

        #endregion constructors

        internal void Modify(Func<XLStyleKey, XLStyleKey> modification)
        {
            Key = modification(Key);

            if (_container != null)
            {
                _container.ModifyStyle(modification);
            }
        }

        #region IXLStyle members

        public IXLFont Font
        {
            get { return new XLFont(this, Value.Font); }
            set
            {
                Modify(k => { k.Font = XLFont.GenerateKey(value); return k; });
            }
        }

        public IXLAlignment Alignment
        {
            get { return new XLAlignment(this, Value.Alignment); }
            set
            {
                Modify(k => { k.Alignment = XLAlignment.GenerateKey(value); return k; });
            }
        }

        public IXLBorder Border
        {
            get { return new XLBorder(_container, this, Value.Border); }
            set
            {
                Modify(k => { k.Border = XLBorder.GenerateKey(value); return k; });
            }
        }

        public IXLFill Fill
        {
            get { return new XLFill(this, Value.Fill); }
            set
            {
                Modify(k => { k.Fill = XLFill.GenerateKey(value); return k; });
            }
        }

        public Boolean IncludeQuotePrefix
        {
            get { return Value.IncludeQuotePrefix; }
            set
            {
                Modify(k => { k.IncludeQuotePrefix = value; return k; });
            }
        }

        public IXLStyle SetIncludeQuotePrefix(Boolean includeQuotePrefix = true)
        {
            IncludeQuotePrefix = includeQuotePrefix;
            return this;
        }

        public IXLNumberFormat NumberFormat
        {
            get { return new XLNumberFormat(this, Value.NumberFormat); }
            set
            {
                Modify(k => { k.NumberFormat = XLNumberFormat.GenerateKey(value); return k; });
            }
        }

        public IXLProtection Protection
        {
            get { return new XLProtection(this, Value.Protection); }
            set
            {
                Modify(k => { k.Protection = XLProtection.GenerateKey(value); return k; });
            }
        }

        public IXLNumberFormat DateFormat
        {
            get { return NumberFormat; }
        }

        #endregion IXLStyle members

        #region Overridden

        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.Append("Font:");
            sb.Append(Font);
            sb.Append(" Fill:");
            sb.Append(Fill);
            sb.Append(" Border:");
            sb.Append(Border);
            sb.Append(" NumberFormat: ");
            sb.Append(NumberFormat);
            sb.Append(" Alignment: ");
            sb.Append(Alignment);
            sb.Append(" Protection: ");
            sb.Append(Protection);
            return sb.ToString();
        }

        public bool Equals(IXLStyle other)
        {
            var otherS = other as XLStyle;

            if (otherS == null)
                return false;

            return Key == otherS.Key &&
                   _container == otherS._container;
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as XLStyle);
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
