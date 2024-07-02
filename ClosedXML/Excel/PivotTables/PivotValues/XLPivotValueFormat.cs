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
            get => _pivotValue.NumberFormatValue?.NumberFormatId ?? -1;
            set
            {
                if (value == -1)
                {
                    _pivotValue.NumberFormatValue = null;
                    return;
                }

                var key = new XLNumberFormatKey
                {
                    NumberFormatId = value,
                    Format = string.Empty,
                };
                _pivotValue.NumberFormatValue = XLNumberFormatValue.FromKey(ref key);
            }
        }

        public String Format
        {
            get => _pivotValue.NumberFormatValue?.Format ?? string.Empty;
            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    _pivotValue.NumberFormatValue = null;
                    return;
                }

                var key = new XLNumberFormatKey
                {
                    NumberFormatId = -1,
                    Format = value,
                };
                _pivotValue.NumberFormatValue = XLNumberFormatValue.FromKey(ref key);
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
            NumberFormatId = value switch
            {
                "General" => 0,
                "0" => 1,
                "0.00" => 2,
                "#,##0" => 3,
                "#,##0.00" => 4,
                "0%" => 9,
                "0.00%" => 10,
                "0.00E+00" => 11,
                _ => -1,
            };

            return _pivotValue;
        }
    }
}
