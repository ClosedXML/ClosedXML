using System;
using System.Collections.Generic;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLCFDataBar : IXLCFDataBar
    {
        private readonly XLConditionalFormat _conditionalFormat;

        public XLCFDataBar(XLConditionalFormat conditionalFormat)
        {
            _conditionalFormat = conditionalFormat;
            _conditionalFormat.DataBar = this;
            // Default value in Excel is true
            Gradient = true;
        }

        public bool Gradient { get; set; }

        public IXLCFDataBar SetGradient(bool value = true)
        {
            Gradient = value;
            return this;
        }
    }
}
