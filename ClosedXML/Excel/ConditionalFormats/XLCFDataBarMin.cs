namespace ClosedXML.Excel
{
    internal class XLCFDataBarMin : IXLCFDataBarMin
    {
        private readonly XLConditionalFormat _conditionalFormat;
        public XLCFDataBarMin(XLConditionalFormat conditionalFormat)
        {
            _conditionalFormat = conditionalFormat;
        }

        public IXLCFDataBarMax Minimum(XLCFContentType type, string value)
        {
            _conditionalFormat.ContentTypes.Initialize(type);
            _conditionalFormat.Values.Initialize(new XLFormula { Value = value });
            return new XLCFDataBarMax(_conditionalFormat);
        }
        public IXLCFDataBarMax Minimum(XLCFContentType type, double value)
        {
            return Minimum(type, value.ToInvariantString());
        }

        public IXLCFDataBarMax LowestValue()
        {
            return Minimum(XLCFContentType.Minimum, "0");
        }
    }
}
