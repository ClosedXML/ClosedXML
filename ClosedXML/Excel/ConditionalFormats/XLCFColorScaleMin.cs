namespace ClosedXML.Excel
{
    internal class XLCFColorScaleMin : IXLCFColorScaleMin
    {
        private readonly XLConditionalFormat _conditionalFormat;
        public XLCFColorScaleMin(XLConditionalFormat conditionalFormat)
        {
            _conditionalFormat = conditionalFormat;
        }

        public IXLCFColorScaleMid Minimum(XLCFContentType type, string value, XLColor color)
        {
            _conditionalFormat.Values.Initialize(new XLFormula { Value = value });
            _conditionalFormat.Colors.Initialize(color);
            _conditionalFormat.ContentTypes.Initialize(type);
            return new XLCFColorScaleMid(_conditionalFormat);
        }
        public IXLCFColorScaleMid Minimum(XLCFContentType type, double value, XLColor color)
        {
            return Minimum(type, value.ToInvariantString(), color);
        }

        public IXLCFColorScaleMid LowestValue(XLColor color)
        {
            _conditionalFormat.Values.Initialize(null);
            _conditionalFormat.Colors.Initialize(color);
            _conditionalFormat.ContentTypes.Initialize(XLCFContentType.Minimum);
            return new XLCFColorScaleMid(_conditionalFormat);
        }
    }
}
