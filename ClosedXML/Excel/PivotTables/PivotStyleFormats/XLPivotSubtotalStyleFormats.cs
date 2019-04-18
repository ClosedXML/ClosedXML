using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal class XLPivotSubtotalStyleFormats : IXLPivotElementStyleFormats
    {
        private readonly List<IXLPivotValueStyleFormat> _dataValuesFormats = new List<IXLPivotValueStyleFormat>();
        private IXLPivotStyleFormat _labelFormat;

        public XLPivotSubtotalStyleFormats(IXLPivotField field)
        {
            PivotField = field;
        }

        public IXLPivotField PivotField { get; }

        public IXLPivotValueStyleFormat AddValuesFormat()
        {
            var dataValuesFormat = new XLPivotValueStyleFormat(PivotField)
            {
                AppliesTo = XLPivotStyleFormatElement.Data,
                Outline = true
            };
            _dataValuesFormats.Add(dataValuesFormat);
            return dataValuesFormat;
        }

        public IEnumerable<IXLPivotValueStyleFormat> DataValuesFormats => _dataValuesFormats;
        public bool HasLabelFormat => _labelFormat != null;

        public IXLPivotStyleFormat Label
        {
            get
            {
                if (_labelFormat == null)
                {
                    _labelFormat = new XLPivotStyleFormat(PivotField)
                    {
                        AppliesTo = XLPivotStyleFormatElement.Label
                    };
                }
                return _labelFormat;
            }
            set { _labelFormat = value; }
        }
    }
}
