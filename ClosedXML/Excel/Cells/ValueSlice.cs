using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    /// <summary>
    /// A slice of a single worksheet for values of a cell.
    /// </summary>
    internal class ValueSlice : ISlice
    {
        private readonly Slice<XLValueSliceContent> _values = new();
        private readonly SharedStringTable _sst;

        internal ValueSlice(SharedStringTable sst)
        {
            _sst = sst;
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
            ref readonly var cellValue = ref _values[point];
            var type = cellValue.Type;
            var value = cellValue.Value;
            return type switch
            {
                XLDataType.Blank => Blank.Value,
                XLDataType.Boolean => value != 0,
                XLDataType.Number => value,
                XLDataType.Text => _sst[(int)value],
                XLDataType.Error => (XLError)value,
                XLDataType.DateTime => XLCellValue.FromSerialDateTime(value),
                XLDataType.TimeSpan => XLCellValue.FromSerialTimeSpan(value),
                _ => throw new ArgumentOutOfRangeException()
            };
        }

        internal void SetCellValue(XLSheetPoint point, XLCellValue cellValue)
        {
            ref readonly var original = ref _values[point];

            double value;
            if (cellValue.Type == XLDataType.Text)
            {
                if (original.Type == XLDataType.Text)
                {
                    // Change references. Increase first and then decrease to have fewer shuffles assigning same value to a cell.
                    var originalStringId = (int)original.Value;
                    value = _sst.IncreaseRef(cellValue.GetText());
                    _sst.DecreaseRef(originalStringId);
                }
                else
                {
                    // The original value wasn't a text -> just increase ref count to a new text
                    value = _sst.IncreaseRef(cellValue.GetText());
                }
            }
            else
            {
                // New value isn't a text
                if (original.Type == XLDataType.Text)
                {
                    // Dereference original text
                    var originalStringId = (int)original.Value;
                    _sst.DecreaseRef(originalStringId);
                }

                if (cellValue.IsUnifiedNumber)
                    value = cellValue.GetUnifiedNumber();
                else if (cellValue.IsBoolean)
                    value = cellValue.GetBoolean() ? 1 : 0;
                else if (cellValue.IsError)
                    value = (int)cellValue.GetError();
                else
                    value = 0; // blank
            }

            var modified = new XLValueSliceContent(value, cellValue.Type, original.ModifiedAtVersion, original.SharedStringId);
            _values.Set(point, in modified);
        }

        internal int GetShareStringId(XLSheetPoint point)
        {
            // This is the public id, separate from real sharedStringId stored in a value
            return _values[point].SharedStringId;
        }

        internal void SetShareStringId(XLSheetPoint point, int sharedStringId)
        {
            ref readonly var original = ref _values[point];
            if (original.SharedStringId != sharedStringId)
            {
                var modified = new XLValueSliceContent(original.Value, original.Type, original.ModifiedAtVersion, sharedStringId);
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
                var modified = new XLValueSliceContent(original.Value, original.Type, modifiedAtVersion, original.SharedStringId);
                _values.Set(point, in modified);
            }
        }

        private readonly struct XLValueSliceContent
        {
            /// <summary>
            /// A cell value in a very compact representation. The value is interpreted depending on a type.
            /// </summary>
            internal readonly double Value;

            /// <summary>
            /// Type of a cell <see cref="Value"/>.
            /// </summary>
            internal readonly XLDataType Type;
            internal readonly long ModifiedAtVersion;
            internal readonly int SharedStringId;

            internal XLValueSliceContent(double value, XLDataType type, long modifiedAtVersion, int sharedStringId)
            {
                Value = value;
                Type = type;
                ModifiedAtVersion = modifiedAtVersion;
                SharedStringId = sharedStringId;
            }
        }
    }
}
