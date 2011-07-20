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
            return _pivotValue;
        }

    }
}
