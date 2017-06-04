using System;

namespace ClosedXML.Excel
{
    internal class XLNumberFormat : IXLNumberFormat
    {
        #region IXLNumberFormat Members

        public bool Equals(IXLNumberFormatBase other)
        {
            return
                _numberFormatId == other.NumberFormatId
                && _format == other.Format
                ;
        }

        #endregion

        private void SetStyleChanged()
        {
            if (_container != null) _container.StyleChanged = true;
        }

        public override bool Equals(object obj)
        {
            return Equals((IXLNumberFormatBase)obj);
        }

        public override int GetHashCode()
        {
            return NumberFormatId
                   ^ Format.GetHashCode();
        }

        #region Properties

        private readonly IXLStylized _container;
        private String _format = String.Empty;

        private Int32 _numberFormatId;

        public Int32 NumberFormatId
        {
            get { return _numberFormatId; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.NumberFormat.NumberFormatId = value);
                else
                {
                    _numberFormatId = value;
                    _format = String.Empty;
                }
            }
        }

        public String Format
        {
            get { return _format; }
            set
            {
                SetStyleChanged();
                if (_container != null && !_container.UpdatingStyle)
                    _container.Styles.ForEach(s => s.NumberFormat.Format = value);
                else
                {
                    _format = value;
                    _numberFormatId = -1;
                }
            }
        }

        public IXLStyle SetNumberFormatId(Int32 value)
        {
            NumberFormatId = value;
            return _container.Style;
        }

        public IXLStyle SetFormat(String value)
        {
            Format = value;
            return _container.Style;
        }

        #endregion

        #region Constructors

        public XLNumberFormat()
            : this(null, XLWorkbook.DefaultStyle.NumberFormat)
        {
        }


        public XLNumberFormat(IXLStylized container, IXLNumberFormat defaultNumberFormat)
        {
            _container = container;
            if (defaultNumberFormat == null) return;
            _numberFormatId = defaultNumberFormat.NumberFormatId;
            _format = defaultNumberFormat.Format;
        }

        #endregion

        #region Overridden

        public override string ToString()
        {
            return _numberFormatId + "-" + _format;
        }

        #endregion
    }
}
