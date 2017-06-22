using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLCFDataBarMax : IXLCFDataBarMax
    {
        private readonly XLConditionalFormat _conditionalFormat;
        public XLCFDataBarMax(XLConditionalFormat conditionalFormat)
        {
            _conditionalFormat = conditionalFormat;
        }

        public void Maximum(XLCFContentType type, String value)
        {
            _conditionalFormat.ContentTypes.Add(type);
            _conditionalFormat.Values.Add(new XLFormula { Value = value });
        }
        public void Maximum(XLCFContentType type, Double value)
        {
            Maximum(type, value.ToInvariantString());
        }

        public void HighestValue()
        {
            Maximum(XLCFContentType.Maximum, "0");
        }
    }
}
