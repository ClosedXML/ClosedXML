namespace ClosedXML.Excel
{
    internal class XLIgnoredError : IXLIgnoredError
    {
        public XLIgnoredError(XLIgnoredErrorType type, IXLRange range)
        {
            Type = type;
            Range = range;
        }
        public XLIgnoredErrorType Type { get; private set; }
        public IXLRange Range { get; private set; }

        public override string ToString()
        {
            return $"{Type}:{Range}";
        }
    }
}
