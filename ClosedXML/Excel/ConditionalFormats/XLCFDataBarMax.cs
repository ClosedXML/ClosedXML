namespace ClosedXML.Excel
{
    internal class XLCFDataBarMax : IXLCFDataBarMax
    {
        private readonly XLConditionalFormat _conditionalFormat;
        public XLCFDataBarMax(XLConditionalFormat conditionalFormat)
        {
            _conditionalFormat = conditionalFormat;
        }

        public void Maximum(XLCFContentType type, string value)
        {
            _conditionalFormat.ContentTypes.Add(type);
            _conditionalFormat.Values.Add(new XLFormula { Value = value });
        }
        public void Maximum(XLCFContentType type, double value)
        {
            Maximum(type, value.ToInvariantString());
        }

        public void HighestValue()
        {
            Maximum(XLCFContentType.Maximum, "0");
        }
    }
}
