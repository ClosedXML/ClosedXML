namespace ClosedXML.Excel
{
    /// <summary>
    /// A slice of a single worksheet for values of a cell.
    /// </summary>
    internal class ValueSlice : Slice<XLValueSliceContent>
    {
        private readonly Slice<XLValueSliceContent> _values;

        internal ValueSlice()
        {
            _values = this;
        }

        internal XLCellValue GetCellValue(XLSheetPoint point)
        {
            return _values[point].Value;
        }

        internal void SetCellValue(XLSheetPoint point, XLCellValue value)
        {
            ref readonly var original = ref _values[point];
            var modified = new XLValueSliceContent(value, original.ModifiedAtVersion, original.SharedStringId);
            _values.Set(point, in modified);
        }

        internal int GetShareStringId(XLSheetPoint point)
        {
            return _values[point].SharedStringId;
        }

        internal void SetShareStringId(XLSheetPoint point, int sharedStringId)
        {
            ref readonly var original = ref _values[point];
            if (original.SharedStringId != sharedStringId)
            {
                var modified = new XLValueSliceContent(original.Value, original.ModifiedAtVersion, sharedStringId);
                _values.Set(point, in modified);
            }
        }

        internal long GetModifiedAtVersion(XLSheetPoint point)
        {
            return _values[point].ModifiedAtVersion;
        }

        internal void SetModifiedAtVersion(XLSheetPoint point, long modifiedAtVersion)
        {
            ref readonly var original = ref _values[point];
            if (original.ModifiedAtVersion != modifiedAtVersion)
            {
                var modified = new XLValueSliceContent(original.Value, modifiedAtVersion, original.SharedStringId);
                _values.Set(point, in modified);
            }
        }
    }
}
