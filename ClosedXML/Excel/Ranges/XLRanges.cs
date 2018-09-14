using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel.Ranges.Index;

namespace ClosedXML.Excel
{
    using System.Collections;

    internal class XLRanges : XLStylizedBase, IXLRanges, IXLStylized
    {
        /// <summary>
        /// Normally, XLRanges collection includes ranges from a single worksheet, but not necessarily.
        /// </summary>
        private readonly Dictionary<IXLWorksheet, IXLRangeIndex<XLRange>> _indexes;
        private IEnumerable<XLRange> Ranges => _indexes.Values.SelectMany(index => index.GetAll());


        private IXLRangeIndex<XLRange> GetRangeIndex(IXLWorksheet worksheet)
        {
            if (!_indexes.ContainsKey(worksheet))
                _indexes.Add(worksheet, new XLRangeIndex<XLRange>(worksheet));

            return _indexes[worksheet];
        }

        public XLRanges() : base(XLWorkbook.DefaultStyleValue)
        {
            _indexes = new Dictionary<IXLWorksheet, IXLRangeIndex<XLRange>>();
        }

        #region IXLRanges Members

        public IXLRanges Clear(XLClearOptions clearOptions = XLClearOptions.All)
        {
            Ranges.ForEach(c => c.Clear(clearOptions));
            return this;
        }

        public void Add(XLRange range)
        {
            if (GetRangeIndex(range.Worksheet).Add(range))
                Count++;
        }

        public void Add(IXLRangeBase range)
        {
            Add(range.AsRange() as XLRange);
        }

        public void Add(IXLCell cell)
        {
            Add(cell.AsRange());
        }

        public void Remove(IXLRange range)
        {
            if (GetRangeIndex(range.Worksheet).Remove(range))
                Count--;
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
            foreach (var index in _indexes.Values)
            {
                Count -= index.RemoveAll(match ?? (_ => true));
            }
        }

        public int Count { get; private set; }

        public IEnumerator<IXLRange> GetEnumerator()
        {
            return Ranges.Cast<IXLRange>().GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public Boolean Contains(IXLCell cell)
        {
            return GetIntersectedRanges((XLAddress)cell.Address).Any();
        }

        public Boolean Contains(IXLRange range)
        {
            return GetIntersectedRanges((XLRangeAddress)range.RangeAddress)
                .Any(r => r.Contains(range));
        }

        /// <summary>
        /// Filter ranges from a collection that intersect the specified address. Is much more efficient
        /// that using Linq expression .Where().
        /// </summary>
        public IEnumerable<IXLRange> GetIntersectedRanges(IXLRangeAddress rangeAddress)
        {
            var xlRangeAddress = (XLRangeAddress)rangeAddress;
            return GetIntersectedRanges(in xlRangeAddress);
        }

        internal IEnumerable<IXLRange> GetIntersectedRanges(in XLRangeAddress rangeAddress)
        {
            return GetRangeIndex(rangeAddress.Worksheet)
                .GetIntersectedRanges(rangeAddress);
        }

        /// <summary>
        /// Filter ranges from a collection that intersect the specified address. Is much more efficient
        /// that using Linq expression .Where().
        /// </summary>
        public IEnumerable<IXLRange> GetIntersectedRanges(IXLAddress address)
        {
            var xlAddress = (XLAddress) address;
            return GetIntersectedRanges(in xlAddress);
        }

        internal IEnumerable<IXLRange> GetIntersectedRanges(in XLAddress address)
        {
            return GetRangeIndex(address.Worksheet)
                .GetIntersectedRanges(address);
        }

        public IEnumerable<IXLRange> GetIntersectedRanges(IXLCell cell)
        {
            return GetIntersectedRanges(cell.Address);
        }

        public IEnumerable<IXLDataValidation> DataValidation
        {
            get { return Ranges.Select(range => range.DataValidation).Where(dv => dv != null); }
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
            Ranges.ForEach(r => r.AddToNamed(rangeName, scope, comment));
            return this;
        }

        public Object Value
        {
            set { Ranges.ForEach(r => r.Value = value); }
        }

        public IXLRanges SetValue<T>(T value)
        {
            Ranges.ForEach(r => r.SetValue(value));
            return this;
        }

        public IXLCells Cells()
        {
            var cells = new XLCells(false, XLCellsUsedOptions.AllContents);
            foreach (XLRange container in Ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLCells CellsUsed()
        {
            var cells = new XLCells(true, XLCellsUsedOptions.AllContents);
            foreach (XLRange container in Ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]

        public IXLCells CellsUsed(Boolean includeFormats)
        {
            return CellsUsed(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents);
        }

        public IXLCells CellsUsed(XLCellsUsedOptions options)
        {
            var cells = new XLCells(true, options);
            foreach (XLRange container in Ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLRanges SetDataType(XLDataType dataType)
        {
            Ranges.ForEach(c => c.DataType = dataType);
            return this;
        }

        #endregion IXLRanges Members

        #region IXLStylized Members

        public override IEnumerable<IXLStyle> Styles
        {
            get
            {
                yield return Style;
                foreach (XLRange rng in Ranges)
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
                foreach (XLRange rng in Ranges)
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
            String retVal = Ranges.Aggregate(String.Empty, (agg, r) => agg + (r.ToString() + ","));
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

            return Ranges.Count() == other.Ranges.Count() &&
                   Ranges.Select(thisRange => Enumerable.Contains(other.Ranges, thisRange)).All(foundOne => foundOne);
        }

        public override int GetHashCode()
        {
            return Ranges.Aggregate(0, (current, r) => current ^ r.GetHashCode());
        }

        public IXLDataValidation SetDataValidation()
        {
            foreach (XLRange range in Ranges)
            {
                foreach (IXLDataValidation dv in range.Worksheet.DataValidations)
                {
                    foreach (IXLRange dvRange in dv.Ranges.GetIntersectedRanges(range.RangeAddress))
                    {
                        dv.Ranges.Remove(dvRange);
                        foreach (IXLCell c in dvRange.Cells().Where(c => !range.Contains(c.Address.ToString())))
                        {
                            dv.Ranges.Add(c.AsRange());
                        }
                    }
                }
            }
            var dataValidation = new XLDataValidation(this);

            Ranges.First().Worksheet.DataValidations.Add(dataValidation);
            return dataValidation;
        }

        public void Select()
        {
            foreach (var range in this)
                range.Select();
        }

        public IXLRanges Consolidate()
        {
            var engine = new XLRangeConsolidationEngine(this);
            return engine.Consolidate();
        }
    }
}
