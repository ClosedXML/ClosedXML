using System;

namespace ClosedXML.Excel
{
    internal class XLProtection : IXLProtection
    {
        readonly IXLStylized _container;

        private Boolean _locked;
        public Boolean Locked
        {
            get
            {
                return _locked;
            }
            set
            {
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Protection.Locked = value);
                else
                    _locked = value;
            }
        }

        private Boolean _hidden;
        public Boolean Hidden
        {
            get
            {
                return _hidden;
            }
            set
            {
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Protection.Hidden = value);
                else
                    _hidden = value;
            }
        }

        #region Constructors

        public XLProtection()
            : this(null, XLWorkbook.DefaultStyle.Protection)
        {
        }

        public XLProtection(IXLStylized container, IXLProtection defaultProtection = null)
        {
            _container = container;
            if (defaultProtection == null) return;

            _locked = defaultProtection.Locked;
            _hidden = defaultProtection.Hidden;
        }

        #endregion

        public bool Equals(IXLProtection other)
        {
            var otherP = other as XLProtection;
            if (otherP == null)
                return false;

            return _locked == otherP._locked
                   && _hidden == otherP._hidden;
        }

        public override bool Equals(object obj)
        {
            return Equals((IXLProtection)obj);
        }

        public override int GetHashCode()
        {
            if (Locked)
                return Hidden ? 11 : 10;

            return Hidden ? 1 : 0;
        }

        public override string ToString()
        {
            if (Locked)
                return Hidden ? "Locked-Hidden" : "Locked";

            return Hidden ? "Hidden" : "None";
        }

        public IXLStyle SetLocked() { Locked = true; return _container.Style; }	public IXLStyle SetLocked(Boolean value) { Locked = value; return _container.Style; }
        public IXLStyle SetHidden() { Hidden = true; return _container.Style; }	public IXLStyle SetHidden(Boolean value) { Hidden = value; return _container.Style; }

    }

}
