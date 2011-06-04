using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLProtection : IXLProtection
    {
        IXLStylized container;

        private Boolean locked;
        public Boolean Locked
        {
            get
            {
                return locked;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Protection.Locked = value);
                else
                    locked = value;
            }
        }

        private Boolean hidden;
        public Boolean Hidden
        {
            get
            {
                return hidden;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Protection.Hidden = value);
                else
                    hidden = value;
            }
        }

        #region Constructors

        public XLProtection()
            : this(null, XLWorkbook.DefaultStyle.Protection)
        {
        }

        public XLProtection(IXLStylized container, IXLProtection defaultProtection = null)
        {
            this.container = container;
            if (defaultProtection != null)
            {
                locked = defaultProtection.Locked;
                hidden = defaultProtection.Hidden;
            }
        }

        #endregion

        public bool Equals(IXLProtection other)
        {
            var otherP = other as XLProtection;
            return locked == otherP.locked
                   && hidden == otherP.hidden;
        }

        public override bool Equals(object obj)
        {
            return this.Equals((IXLProtection)obj);
        }

        public override int GetHashCode()
        {
            if (Locked)
                if (Hidden)
                    return 11;
                else
                    return 10;
            else
                if (Hidden)
                    return 1;
                else
                    return 0;
        }

        public override string ToString()
        {
            if (Locked)
                if (Hidden)
                    return "Locked-Hidden";
                else
                    return "Locked";
            else
                if (Hidden)
                    return "Hidden";
                else
                    return "None";
        }

        public IXLStyle SetLocked() { Locked = true; return container.Style; }	public IXLStyle SetLocked(Boolean value) { Locked = value; return container.Style; }
        public IXLStyle SetHidden() { Hidden = true; return container.Style; }	public IXLStyle SetHidden(Boolean value) { Hidden = value; return container.Style; }

    }

}
