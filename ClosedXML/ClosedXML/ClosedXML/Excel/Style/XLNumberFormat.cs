using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLNumberFormat : IXLNumberFormat
    {
        #region Properties

        IXLStylized container;

        private Int32 numberFormatId;
        public Int32 NumberFormatId
        {
            get { return numberFormatId; }
            set
            {
                if (container != null && !container.UpdatingStyle)
                {
                    container.Styles.ForEach(s => s.NumberFormat.NumberFormatId = value);
                }
                else
                {
                    numberFormatId = value;
                    format = String.Empty;
                }
            }
        }

        private String format = String.Empty;
        public String Format
        {
            get { return format; }
            set
            {
                if (container != null && !container.UpdatingStyle)
                {
                    container.Styles.ForEach(s => s.NumberFormat.Format = value);
                }
                else
                {
                    format = value;
                    numberFormatId = -1;
                }
            }
        }

        #endregion

        #region Constructors

        public XLNumberFormat()
            : this(null, XLWorkbook.DefaultStyle.NumberFormat)
        {
        }


        public XLNumberFormat(IXLStylized container, IXLNumberFormat defaultNumberFormat)
        {
            this.container = container;
            if (defaultNumberFormat != null)
            {
                numberFormatId = defaultNumberFormat.NumberFormatId;
                format = defaultNumberFormat.Format;
            }
        }

        #endregion

        #region Overridden

        public override string ToString()
        {
            return numberFormatId.ToString() + "-" + format;
        }

        #endregion

        public bool Equals(IXLNumberFormat other)
        {
            return
            this.NumberFormatId.Equals(other.NumberFormatId)
            && this.Format.Equals(other.Format)
            ;
        }

        public override bool Equals(object obj)
        {
            return this.Equals((XLNumberFormat)obj);
        }

        public override int GetHashCode()
        {
            unchecked // Overflow is fine, just wrap
            {
                int hash = 17;
                hash = hash * 23 + NumberFormatId.GetHashCode();
                hash = hash * 23 + Format.GetHashCode();
                return hash;
            }
        }
    }
}
