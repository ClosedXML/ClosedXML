// Keep this file CodeMaid organised and cleaned

namespace ClosedXML.Excel
{
    internal class XLPivotFieldStyleFormats : IXLPivotFieldStyleFormats
    {
        private IXLPivotValueStyleFormat dataValuesFormat;
        private IXLPivotStyleFormat headerFormat;
        private IXLPivotStyleFormat labelFormat;
        private IXLPivotStyleFormat subtotalFormat;

        public XLPivotFieldStyleFormats(IXLPivotField field)
        {
            this.PivotField = field;
        }

        public IXLPivotField PivotField { get; }

        #region IXLPivotFieldStyleFormats

        public IXLPivotValueStyleFormat DataValuesFormat
        {
            get
            {
                if (dataValuesFormat == null)
                {
                    dataValuesFormat = new XLPivotValueStyleFormat(PivotField)
                    {
                        AppliesTo = XLPivotStyleFormatElement.Data
                    };
                }
                return dataValuesFormat;
            }
            set { dataValuesFormat = value; }
        }

        public IXLPivotStyleFormat Header
        {
            get
            {
                if (headerFormat == null)
                {
                    headerFormat = new XLPivotStyleFormat(PivotField);
                }
                return headerFormat;
            }
            set { headerFormat = value; }
        }

        public IXLPivotStyleFormat Label
        {
            get
            {
                if (labelFormat == null)
                {
                    labelFormat = new XLPivotStyleFormat(PivotField)
                    {
                        AppliesTo = XLPivotStyleFormatElement.Label
                    };
                }
                return labelFormat;
            }
            set { labelFormat = value; }
        }

        public IXLPivotStyleFormat Subtotal
        {
            get
            {
                if (subtotalFormat == null)
                {
                    subtotalFormat = new XLPivotStyleFormat(PivotField);
                }

                return subtotalFormat;
            }
            set { subtotalFormat = value; }
        }

        #endregion IXLPivotFieldStyleFormats
    }
}
