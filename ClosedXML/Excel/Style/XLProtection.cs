using System;

namespace ClosedXML.Excel
{
    internal class XLProtection : IXLProtection
    {
        private readonly IXLStylized _container;
        private Boolean _hidden;

        private Boolean _locked;

        #region IXLProtection Members

        public Boolean Locked
        {
            get { return _locked; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Protection.Locked = value);
                else
                    _locked = value;
            }
        }

        public Boolean Hidden
        {
            get { return _hidden; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.Protection.Hidden = value);
                else
                    _hidden = value;
            }
        }

        public bool Equals(IXLProtection other)
        {
            var otherP = other as XLProtection;
            if (otherP == null)
                return false;

            return _locked == otherP._locked
                   && _hidden == otherP._hidden;
        }

        public IXLStyle SetLocked()
        {
            Locked = true;
            return _container.Style;
        }

        public IXLStyle SetLocked(Boolean value)
        {
            Locked = value;
            return _container.Style;
        }

        public IXLStyle SetHidden()
        {
            Hidden = true;
            return _container.Style;
        }

        public IXLStyle SetHidden(Boolean value)
        {
            Hidden = value;
            return _container.Style;
        }

        #endregion

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

        private void SetStyleChanged()
        {
            if (_container != null) _container.StyleChanged = true;
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
    }
}