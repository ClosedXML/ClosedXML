using System;

namespace ClosedXML.Excel
{
    internal class XLCFColorScaleMid : IXLCFColorScaleMid
    {
        private readonly XLConditionalFormat _conditionalFormat;
        public XLCFColorScaleMid(XLConditionalFormat conditionalFormat)
        {
            _conditionalFormat = conditionalFormat;
        }
        public IXLCFColorScaleMax Midpoint(XLCFContentType type, String value, XLColor color)
        {
            _conditionalFormat.Values.Add(new XLFormula { Value = value });
            _conditionalFormat.Colors.Add(color);
            _conditionalFormat.ContentTypes.Add(type);
            return new XLCFColorScaleMax(_conditionalFormat);
        }
        public IXLCFColorScaleMax Midpoint(XLCFContentType type, Double value, XLColor color)
        {
            return Midpoint(type, value.ToInvariantString(), color);
        }
        public void Maximum(XLCFContentType type, String value, XLColor color)
        {
            Midpoint(type, value, color);
        }
        public void Maximum(XLCFContentType type, Double value, XLColor color)
        {
            Maximum(type, value.ToInvariantString(), color);
        }
        public void HighestValue(XLColor color)
        {
            _conditionalFormat.Values.Initialize(null);
            _conditionalFormat.Colors.Add(color);
            _conditionalFormat.ContentTypes.Add(XLCFContentType.Maximum);
        }
    }
}
