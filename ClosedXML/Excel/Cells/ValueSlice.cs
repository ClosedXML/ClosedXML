using System.Collections.Generic;

namespace ClosedXML.Excel
{
    /// <summary>
    /// A slice of a single worksheet for values of a cell.
    /// </summary>
    internal class ValueSlice : ISlice
    {
        private readonly Slice<XLValueSliceContent> _values;

        internal ValueSlice()
        {
            _values = new();
        }

        public bool IsEmpty => _values.IsEmpty;

        public int MaxColumn => _values.MaxColumn;

        public int MaxRow => _values.MaxRow;

        public Dictionary<int, int>.KeyCollection UsedColumns => _values.UsedColumns;

        public IEnumerable<int> UsedRows => _values.UsedRows;

        public void Clear(XLSheetRange range) => _values.Clear(range);

        public void DeleteAreaAndShiftLeft(XLSheetRange rangeToDelete) => _values.DeleteAreaAndShiftLeft(rangeToDelete);

        public void DeleteAreaAndShiftUp(XLSheetRange rangeToDelete) => _values.DeleteAreaAndShiftUp(rangeToDelete);

        public IEnumerator<XLSheetPoint> GetEnumerator(XLSheetRange range, bool reverse = false) => _values.GetEnumerator(range, reverse);

        public void InsertAreaAndShiftDown(XLSheetRange range) => _values.InsertAreaAndShiftDown(range);

        public void InsertAreaAndShiftRight(XLSheetRange range) => _values.InsertAreaAndShiftRight(range);

        public bool IsUsed(XLSheetPoint address) => _values.IsUsed(address);

        public void Swap(XLSheetPoint sp1, XLSheetPoint sp2) => _values.Swap(sp1, sp2);

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

        private readonly struct XLValueSliceContent
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
}
