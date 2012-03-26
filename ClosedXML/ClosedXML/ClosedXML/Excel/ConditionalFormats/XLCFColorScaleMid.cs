using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLCFColorScaleMid : IXLCFColorScaleMid
    {
        private readonly XLConditionalFormat _conditionalFormat;
        public XLCFColorScaleMid(XLConditionalFormat conditionalFormat)
        {
            _conditionalFormat = conditionalFormat;
        }
        public IXLCFColorScaleMax Midpoint(XLCFContentType type, String value, IXLColor color)
        {
            _conditionalFormat.Values.Add(value);
            _conditionalFormat.Colors.Add(color);
            _conditionalFormat.ContentTypes.Add(type);
            return new XLCFColorScaleMax(_conditionalFormat);
        }
        public void Maximum(XLCFContentType type, String value, IXLColor color)
        {
            Midpoint(type, value, color);
        }
        public void HighestValue(IXLColor color)
        {
            Midpoint(XLCFContentType.Maximum, "0", color);
        }
    }
}
