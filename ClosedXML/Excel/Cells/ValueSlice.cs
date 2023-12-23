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

        public void Clear(XLSheetRange range)
        {
            DereferenceTextInRange(range);
            _values.Clear(range);
        }

        public void DeleteAreaAndShiftLeft(XLSheetRange rangeToDelete)
        {
            DereferenceTextInRange(rangeToDelete);
            _values.DeleteAreaAndShiftLeft(rangeToDelete);
        }

        public void DeleteAreaAndShiftUp(XLSheetRange rangeToDelete)
        {
            DereferenceTextInRange(rangeToDelete);
            _values.DeleteAreaAndShiftUp(rangeToDelete);
        }

        public IEnumerator<XLSheetPoint> GetEnumerator(XLSheetRange range, bool reverse = false) => _values.GetEnumerator(range, reverse);

        public void InsertAreaAndShiftDown(XLSheetRange range)
        {
            // Only pushed out references have to be dereferenced, other text references just move.
            if (range.BottomRow < XLHelper.MaxRowNumber)
            {
                var belowRange = range.BelowRange();
                var pushedOutRows = Math.Min(range.Height, belowRange.Height);
                var pushedOutRange = belowRange.SliceFromBottom(pushedOutRows);
                DereferenceTextInRange(pushedOutRange);
            }

            _values.InsertAreaAndShiftDown(range);
        }

        public void InsertAreaAndShiftRight(XLSheetRange range)
        {
            // Only pushed out references have to be dereferenced, other text references just move.
            if (range.RightColumn < XLHelper.MaxColumnNumber)
            {
                var rightRange = range.RightRange();
                var pushedOutColumns = Math.Min(range.Width, rightRange.Width);
                var pushedOutRange = rightRange.SliceFromRight(pushedOutColumns);
                DereferenceTextInRange(pushedOutRange);
            }

            _values.InsertAreaAndShiftRight(range);
        }

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
                    value = _sst.IncreaseRef(cellValue.GetText(), original.Inline);
                    _sst.DecreaseRef(originalStringId);
                }
                else
                {
                    // The original value wasn't a text -> just increase ref count to a new text
                    value = _sst.IncreaseRef(cellValue.GetText(), original.Inline);
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

            var modified = new XLValueSliceContent(value, cellValue.Type, original.Inline);
            _values.Set(point, in modified);
        }

        internal XLImmutableRichText? GetRichText(XLSheetPoint point)
        {
            ref readonly var cellValue = ref _values[point];
            if (cellValue.Type != XLDataType.Text)
                return null;

            var value = cellValue.Value;
            return _sst.GetRichText((int)value);
        }

        internal void SetRichText(XLSheetPoint point, XLImmutableRichText richText)
        {
            if (richText is null)
                throw new ArgumentNullException(nameof(richText));

            ref readonly var original = ref _values[point];

            // If original value was a text (no matter if plain or rich text),
            // dereference because it's being replaced.
            if (original.Type == XLDataType.Text)
            {
                var originalId = (int)original.Value;
                _sst.DecreaseRef(originalId);
            }

            var richTextId = _sst.IncreaseRef(richText, original.Inline);
            var modified = new XLValueSliceContent(richTextId, XLDataType.Text, original.Inline);
            _values.Set(point, modified);
        }

        internal bool GetShareString(XLSheetPoint point)
        {
            return !_values[point].Inline;
        }

        internal void SetShareString(XLSheetPoint point, bool shareString)
        {
            var inlineString = !shareString;
            ref readonly var original = ref _values[point];
            if (original.Inline == inlineString)
                return;

            var cellValue = original.Value;
            if (original.Type == XLDataType.Text)
            {
                // Because inline is a part of SST, we have to update stringIds when inline flag changes.
                var originalStringId = (int)cellValue;
                var richText = _sst.GetRichText(originalStringId);
                if (richText is not null)
                {
                    // Cell is storing rich text
                    _sst.DecreaseRef(originalStringId);
                    cellValue = _sst.IncreaseRef(richText, inlineString);
                }
                else
                {
                    // Cell is storing plain text.
                    var originalString = _sst[originalStringId];
                    _sst.DecreaseRef(originalStringId);
                    cellValue = _sst.IncreaseRef(originalString, inlineString);
                }
            }

            var modified = new XLValueSliceContent(cellValue, original.Type, inlineString);
            _values.Set(point, in modified);
        }

        internal int GetShareStringId(XLSheetPoint point)
        {
            ref readonly var value = ref _values[point];
            if (value.Type != XLDataType.Text)
                throw new InvalidOperationException($"Asking for a shared string id of a non-text cell {point}.");

            return (int)_values[point].Value;
        }

        /// <summary>
        /// Prepare for worksheet removal, dereference all tests in a slice.
        /// </summary>
        internal void DereferenceSlice() => DereferenceTextInRange(XLSheetRange.Full);

        private void DereferenceTextInRange(XLSheetRange range)
        {
            // Dereference all texts in the range, so the ref count is kept correct.
            using var e = _values.GetEnumerator(range);
            while (e.MoveNext())
            {
                ref readonly var value = ref _values[e.Current];
                if (value.Type == XLDataType.Text)
                {
                    _sst.DecreaseRef((int)value.Value);
                    var blank = new XLValueSliceContent(0, XLDataType.Blank, value.Inline);
                    _values.Set(e.Current, in blank);
                }
            }
        }

        private readonly record struct XLValueSliceContent
        {
            /// <summary>
            /// A cell value in a very compact representation. The value is interpreted depending on a type.
            /// </summary>
            internal readonly double Value;

            /// <summary>
            /// Type of a cell <see cref="Value"/>.
            /// </summary>
            internal readonly XLDataType Type;
            internal readonly bool Inline;

            internal XLValueSliceContent(double value, XLDataType type, bool inline)
            {
                Value = value;
                Type = type;
                Inline = inline;
            }
        }
    }
}
