using System;

namespace ClosedXML.Excel
{
    internal class XLPivotValueFormat : IXLPivotValueFormat
    {
        private readonly XLPivotDataField _pivotValue;

        public XLPivotValueFormat(XLPivotDataField pivotValue)
        {
            _pivotValue = pivotValue;
        }

        public Int32 NumberFormatId
        {
            get => _pivotValue.NumberFormatId ?? -1;
            set
            {
                _pivotValue.NumberFormatId = value == -1 ? null : value;
                _pivotValue.NumberFormatCode = string.Empty;
            }
        }

        public String Format
        {
            get => _pivotValue.NumberFormatCode;
            set
            {
                _pivotValue.NumberFormatCode = value;
                _pivotValue.NumberFormatId = null;
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
            _pivotValue.NumberFormatId = value switch
            {
                "General" => 0,
                "0" => 1,
                "0.00" => 2,
                "#,##0" => 3,
                "#,##0.00" => 4,
                "0%" => 9,
                "0.00%" => 10,
                "0.00E+00" => 11,
                _ => null,
            };

            return _pivotValue;
        }
    }
}
