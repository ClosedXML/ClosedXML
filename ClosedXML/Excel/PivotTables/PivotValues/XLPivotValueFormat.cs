using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLPivotValueFormat: IXLPivotValueFormat
    {
        private readonly XLPivotValue _pivotValue;
        public XLPivotValueFormat(XLPivotValue pivotValue)
        {
            _pivotValue = pivotValue;
            _format = "General";
            _numberFormatId = 0;
        }

        private Int32 _numberFormatId = -1;
        public Int32 NumberFormatId
        {
            get { return _numberFormatId; }
            set
            {
                _numberFormatId = value;
                _format = string.Empty;
            }
        }

        private String _format = String.Empty;
        public String Format
        {
            get { return _format; }
            set
            {
                _format = value;
                _numberFormatId = -1;
            }
        }

        public IXLPivotValue SetNumberFormatId(Int32 value)
        {
            NumberFormatId = value;
            return _pivotValue;
        }
        public IXLPivotValue SetFormat(String value)
        {
            Format = value;

            switch (value)
            {
                case "General":
                    _numberFormatId = 0;
                    break;
                case "0":
                    _numberFormatId = 1;
                    break;
                case "0.00":
                    _numberFormatId = 2;
                    break;
                case "#,##0":
                    _numberFormatId = 3;
                    break;
                case "#,##0.00":
                    _numberFormatId = 4;
                    break;
                case "0%":
                    _numberFormatId = 9;
                    break;
                case "0.00%":
                    _numberFormatId = 10;
                    break;
                case "0.00E+00":
                    _numberFormatId = 11;
                    break;
            }


            return _pivotValue;
        }

        #region Overrides
        public bool Equals(IXLNumberFormatBase other)
        {
            return
                _numberFormatId == other.NumberFormatId
                && _format == other.Format
                ;
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

        #endregion

    }
}
