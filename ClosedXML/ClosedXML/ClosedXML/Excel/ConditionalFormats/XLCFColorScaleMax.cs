using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLCFColorScaleMax : IXLCFColorScaleMax
    {
        private readonly XLConditionalFormat _conditionalFormat;
        public XLCFColorScaleMax(XLConditionalFormat conditionalFormat)
        {
            _conditionalFormat = conditionalFormat;
        }

        public void Maximum(XLCFContentType type, String value, IXLColor color)
        {
            _conditionalFormat.Values.Add(value);
            _conditionalFormat.Colors.Add(color);
            _conditionalFormat.ContentTypes.Add(type);
        }
        public void Maximum(XLCFContentType type, Double value, IXLColor color)
        {
            Maximum(type, value.ToString(), color);
        }
        public void HighestValue(IXLColor color)
        {
            Maximum(XLCFContentType.Maximum, "0", color);
        }
    }
}
