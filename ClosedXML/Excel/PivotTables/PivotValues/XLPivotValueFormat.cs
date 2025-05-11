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
                if (!XLPredefinedFormat.FormatCodes.TryGetValue(value, out var format))
                    throw new ArgumentOutOfRangeException($"Only predefined format is permitted. Use nested enums/members of {nameof(XLPredefinedFormat)}.");

                var key = new XLNumberFormatKey
                {
                    NumberFormatId = value,
                    Format = format
                };
                _pivotValue.NumberFormatValue = XLNumberFormatValue.FromKey(ref key);
            }
        }

        public String Format
        {
            get => _pivotValue.NumberFormatValue?.Format ?? string.Empty;
            set
            {
                var key = XLNumberFormatKey.ForFormat(value);
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
            return _pivotValue;
        }
    }
}
