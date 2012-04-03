using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLCFColorScaleMin : IXLCFColorScaleMin
    {
        private readonly XLConditionalFormat _conditionalFormat;
        public XLCFColorScaleMin(XLConditionalFormat conditionalFormat)
        {
            _conditionalFormat = conditionalFormat;
        }

        public IXLCFColorScaleMid Minimum(XLCFContentType type, String value, IXLColor color)
        {
            _conditionalFormat.Values.Initialize(value);
            _conditionalFormat.Colors.Initialize(color);
            _conditionalFormat.ContentTypes.Initialize(type);
            return new XLCFColorScaleMid(_conditionalFormat);
        }
        public IXLCFColorScaleMid Minimum(XLCFContentType type, Double value, IXLColor color)
        {
            return Minimum(type, value.ToString(), color);
        }

        public IXLCFColorScaleMid LowestValue(IXLColor color)
        {
            return Minimum(XLCFContentType.Minimum, "0", color);
        }
    }
}
