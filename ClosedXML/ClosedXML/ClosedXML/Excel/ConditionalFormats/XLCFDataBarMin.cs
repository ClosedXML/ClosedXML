using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLCFDataBarMin : IXLCFDataBarMin
    {
        private readonly XLConditionalFormat _conditionalFormat;
        public XLCFDataBarMin(XLConditionalFormat conditionalFormat)
        {
            _conditionalFormat = conditionalFormat;
        }

        public IXLCFDataBarMax Minimum(XLCFContentType type, String value)
        {
            _conditionalFormat.ContentTypes.Initialize(type);
            _conditionalFormat.Values.Initialize(value);
            return new XLCFDataBarMax(_conditionalFormat);
        }

        public IXLCFDataBarMax LowestValue()
        {
            return Minimum(XLCFContentType.Minimum, "0");
        }
    }
}
