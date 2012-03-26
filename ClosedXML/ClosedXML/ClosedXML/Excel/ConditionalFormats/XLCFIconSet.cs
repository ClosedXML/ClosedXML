using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLCFIconSet : IXLCFIconSet
    {
        private readonly XLConditionalFormat _conditionalFormat;
        public XLCFIconSet(XLConditionalFormat conditionalFormat)
        {
            _conditionalFormat = conditionalFormat;
        }
        public IXLCFIconSet AddValue(XLCFIconSetOperator setOperator, String value, XLCFContentType type)
        {
            _conditionalFormat.IconSetOperators.Add(setOperator);
            _conditionalFormat.Values.Add(value);
            _conditionalFormat.ContentTypes.Add(type);
            return new XLCFIconSet(_conditionalFormat);
        }
    }
}
