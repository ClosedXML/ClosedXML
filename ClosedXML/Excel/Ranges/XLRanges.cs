using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    using System.Collections;

    internal class XLRanges : XLStylizedBase, IXLRanges, IXLStylized
    {
        private readonly List<XLRange> _ranges = new List<XLRange>();

        public XLRanges() : base(XLWorkbook.DefaultStyleValue)
        {
        }

        #region IXLRanges Members

        public IXLRanges Clear(XLClearOptions clearOptions = XLClearOptions.All)
        {
            _ranges.ForEach(c => c.Clear(clearOptions));
            return this;
        }

        public void Add(XLRange range)
        {
            Count++;
            _ranges.Add(range);
        }

        public void Add(IXLRangeBase range)
        {
            Count++;
            _ranges.Add(range.AsRange() as XLRange);
        }

        public void Add(IXLCell cell)
        {
            Add(cell.AsRange());
        }

        public void Remove(IXLRange range)
        {
            Count--;
            _ranges.RemoveAll(r => r.ToString() == range.ToString());
        }

        /// <summary>
        /// Removes ranges matching the criteria from the collection, optionally releasing their event handlers.
        /// </summary>
        /// <param name="match">Criteria to filter ranges. Only those ranges that satisfy the criteria will be removed.
        /// Null means the entire collection should be cleared.</param>
        /// <param name="releaseEventHandlers">Specify whether or not should removed ranges be unsubscribed from 
        /// row/column shifting events. Until ranges are unsubscribed they cannot be collected by GC.</param>
        public void RemoveAll(Predicate<IXLRange> match = null, bool releaseEventHandlers = true)
        {
            match = match ?? (_ => true);

            if (releaseEventHandlers)
            {
                _ranges
                    .Where(r => match(r))
                    .ForEach(r => r.Dispose());
            }

            Count -= _ranges.RemoveAll(match);
        }

        public int Count { get; private set; }

        public IEnumerator<IXLRange> GetEnumerator()
        {
            var retList = new List<IXLRange>();
            retList.AddRange(_ranges.Where(r => XLHelper.IsValidRangeAddress(r.RangeAddress)).Cast<IXLRange>());
            return retList.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public Boolean Contains(IXLCell cell)
        {
            return _ranges.Any(r => r.RangeAddress.IsValid && r.Contains(cell));
        }

        public Boolean Contains(IXLRange range)
        {
            return _ranges.Any(r => r.RangeAddress.IsValid && r.Contains(range));
        }

        public IEnumerable<IXLDataValidation> DataValidation
        {
            get { return _ranges.Select(range => range.DataValidation).Where(dv => dv != null); }
        }

        public IXLRanges AddToNamed(String rangeName)
        {
            return AddToNamed(rangeName, XLScope.Workbook);
        }

        public IXLRanges AddToNamed(String rangeName, XLScope scope)
        {
            return AddToNamed(rangeName, XLScope.Workbook, null);
        }

        public IXLRanges AddToNamed(String rangeName, XLScope scope, String comment)
        {
            _ranges.ForEach(r => r.AddToNamed(rangeName, scope, comment));
            return this;
        }

        public Object Value
        {
            set { _ranges.ForEach(r => r.Value = value); }
        }

        public IXLRanges SetValue<T>(T value)
        {
            _ranges.ForEach(r => r.SetValue(value));
            return this;
        }

        public IXLCells Cells()
        {
            var cells = new XLCells(false, false);
            foreach (XLRange container in _ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLCells CellsUsed()
        {
            var cells = new XLCells(true, false);
            foreach (XLRange container in _ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLCells CellsUsed(Boolean includeFormats)
        {
            var cells = new XLCells(true, includeFormats);
            foreach (XLRange container in _ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLRanges SetDataType(XLDataType dataType)
        {
            _ranges.ForEach(c => c.DataType = dataType);
            return this;
        }

        public void Dispose()
        {
            _ranges.ForEach(r => r.Dispose());
        }

        #endregion IXLRanges Members

        #region IXLStylized Members

        public override IEnumerable<IXLStyle> Styles
        {
            get
            {
                yield return Style;
                foreach (XLRange rng in _ranges)
                {
                    yield return rng.Style;
                    foreach (XLCell r in rng.Worksheet.Internals.CellsCollection.GetCells(
                        rng.RangeAddress.FirstAddress.RowNumber,
                        rng.RangeAddress.FirstAddress.ColumnNumber,
                        rng.RangeAddress.LastAddress.RowNumber,
                        rng.RangeAddress.LastAddress.ColumnNumber))
                        yield return r.Style;
                }
            }
        }

        protected override IEnumerable<XLStylizedBase> Children
        {
            get
            {
                foreach (XLRange rng in _ranges)
                    yield return rng;
            }
        }

        public override IXLRanges RangesUsed
        {
            get { return this; }
        }

        #endregion IXLStylized Members

        public override string ToString()
        {
            String retVal = _ranges.Aggregate(String.Empty, (agg, r) => agg + (r.ToString() + ","));
            if (retVal.Length > 0) retVal = retVal.Substring(0, retVal.Length - 1);
            return retVal;
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as XLRanges);
        }

        public bool Equals(XLRanges other)
        {
            if (other == null)
                return false;

            return _ranges.Count == other._ranges.Count &&
                   _ranges.Select(thisRange => Enumerable.Contains(other._ranges, thisRange)).All(foundOne => foundOne);
        }

        public override int GetHashCode()
        {
            return _ranges.Aggregate(0, (current, r) => current ^ r.GetHashCode());
        }

        public IXLDataValidation SetDataValidation()
        {
            foreach (XLRange range in _ranges)
            {
                foreach (IXLDataValidation dv in range.Worksheet.DataValidations)
                {
                    foreach (IXLRange dvRange in dv.Ranges.Where(dvRange => dvRange.Intersects(range)))
                    {
                        dv.Ranges.Remove(dvRange);
                        foreach (IXLCell c in dvRange.Cells().Where(c => !range.Contains(c.Address.ToString())))
                        {
                            var r = c.AsRange();
                            r.Dispose();
                            dv.Ranges.Add(r);
                        }
                    }
                }
            }
            var dataValidation = new XLDataValidation(this);

            _ranges.First().Worksheet.DataValidations.Add(dataValidation);
            return dataValidation;
        }

        public void Select()
        {
            foreach (var range in this)
                range.Select();
        }
    }
}
