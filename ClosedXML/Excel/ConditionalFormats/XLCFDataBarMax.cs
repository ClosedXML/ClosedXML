#nullable disable

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

        public IXLCFDataBar Maximum(XLCFContentType type, String value)
        {
            _conditionalFormat.ContentTypes.Add(type);
            _conditionalFormat.Values.Add(new XLFormula { Value = value });
            return new XLCFDataBar(_conditionalFormat);
        }

        public IXLCFDataBar Maximum(XLCFContentType type, Double value)
        {
            return Maximum(type, value.ToInvariantString());
        }

        public IXLCFDataBar HighestValue()
        {
            return Maximum(XLCFContentType.Maximum, "0");
        }
    }
}
