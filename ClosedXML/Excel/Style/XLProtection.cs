using System;

namespace ClosedXML.Excel
{
    internal class XLProtection : IXLProtection
    {
        #region Static members

        internal static XLProtectionKey GenerateKey(IXLProtection defaultProtection)
        {
            if (defaultProtection == null)
                return XLProtectionValue.Default.Key;
            if (defaultProtection is XLProtection)
                return (defaultProtection as XLProtection).Key;

            return new XLProtectionKey
            {
                Locked = defaultProtection.Locked,
                Hidden = defaultProtection.Hidden
            };
        }

        #endregion Static members

        #region Properties

        private readonly XLStyle _style;

        private XLProtectionValue _value;

        internal XLProtectionKey Key
        {
            get { return _value.Key; }
            private set { _value = XLProtectionValue.FromKey(ref value); }
        }

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Create an instance of XLProtection initializing it with the specified value.
        /// </summary>
        /// <param name="style">Style to attach the new instance to.</param>
        /// <param name="value">Style value to use.</param>
        public XLProtection(XLStyle style, XLProtectionValue value)
        {
            _style = style ?? XLStyle.CreateEmptyStyle();
            _value = value;
        }

        public XLProtection(XLStyle style, XLProtectionKey key) : this(style, XLProtectionValue.FromKey(ref key))
        {
        }

        public XLProtection(XLStyle style = null, IXLProtection d = null) : this(style, GenerateKey(d))
        {
        }

        #endregion Constructors

        #region IXLProtection Members

        public Boolean Locked
        {
            get { return Key.Locked; }
            set
            {
                Modify(k => { k.Locked = value; return k; });
            }
        }

        public Boolean Hidden
        {
            get { return Key.Hidden; }
            set
            {
                Modify(k => { k.Hidden = value; return k; });
            }
        }

        public IXLStyle SetLocked()
        {
            Locked = true;
            return _style;
        }

        public IXLStyle SetLocked(Boolean value)
        {
            Locked = value;
            return _style;
        }

        public IXLStyle SetHidden()
        {
            Hidden = true;
            return _style;
        }

        public IXLStyle SetHidden(Boolean value)
        {
            Hidden = value;
            return _style;
        }

        #endregion IXLProtection Members

        private void Modify(Func<XLProtectionKey, XLProtectionKey> modification)
        {
            Key = modification(Key);

            _style.Modify(styleKey =>
            {
                var protection = styleKey.Protection;
                styleKey.Protection = modification(protection);
                return styleKey;
            });
        }

        #region Overridden

        public override bool Equals(object obj)
        {
            return Equals((IXLProtection)obj);
        }

        public bool Equals(IXLProtection other)
        {
            var otherP = other as XLProtection;
            if (otherP == null)
                return false;

            return Key == otherP.Key;
        }

        public override string ToString()
        {
            if (Locked)
                return Hidden ? "Locked-Hidden" : "Locked";

            return Hidden ? "Hidden" : "None";
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
