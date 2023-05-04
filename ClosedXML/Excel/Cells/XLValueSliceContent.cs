namespace ClosedXML.Excel
{
    internal readonly struct XLValueSliceContent
    {
        public readonly XLCellValue Value;
        public readonly long ModifiedAtVersion;
        public readonly int SharedStringId;

        public XLValueSliceContent(XLCellValue value, long modifiedAtVersion, int sharedStringId)
        {
            Value = value;
            ModifiedAtVersion = modifiedAtVersion;
            SharedStringId = sharedStringId;
        }
    }
}
