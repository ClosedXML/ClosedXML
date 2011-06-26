using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLCells : IXLCells, IXLStylized, IEnumerable<XLCell>
    {
        #region Fields

        private readonly bool _includeFormats;
        private readonly List<XLRangeAddress> _rangeAddresses = new List<XLRangeAddress>();
        private readonly bool _usedCellsOnly;
        private IXLStyle _style;

        #endregion

        #region Constructor

        public XLCells(bool usedCellsOnly, bool includeFormats)
        {
            _style = new XLStyle(this, XLWorkbook.DefaultStyle);
            _usedCellsOnly = usedCellsOnly;
            _includeFormats = includeFormats;
        }

        #endregion

        #region IEnumerable<XLCell> Members

        public IEnumerator<XLCell> GetEnumerator()
        {
            var cellsInRanges = new Dictionary<XLWorksheet, HashSet<IXLAddress>>();
            foreach (XLRangeAddress range in _rangeAddresses)
            {
                HashSet<IXLAddress> hash;
                if (cellsInRanges.ContainsKey(range.Worksheet))
                    hash = cellsInRanges[range.Worksheet];
                else
                {
                    hash = new HashSet<IXLAddress>();
                    cellsInRanges.Add(range.Worksheet, hash);
                }

                if (_usedCellsOnly)
                {
                    var tmpRange = range;
                    var addressList = range.Worksheet.Internals.CellsCollection.Keys
                        .Where(a => a.RowNumber >= tmpRange.FirstAddress.RowNumber &&
                                    a.RowNumber <= tmpRange.LastAddress.RowNumber &&
                                    a.ColumnNumber >= tmpRange.FirstAddress.ColumnNumber &&
                                    a.ColumnNumber <= tmpRange.LastAddress.ColumnNumber);

                    foreach (IXLAddress a in addressList)
                    {
                        if (!hash.Contains(a))
                            hash.Add(a);
                    }
                }
                else
                {
                    var mm = new MinMax
                                 {
                                     MinRow = range.FirstAddress.RowNumber,
                                     MaxRow = range.LastAddress.RowNumber,
                                     MinColumn = range.FirstAddress.ColumnNumber,
                                     MaxColumn = range.LastAddress.ColumnNumber
                                 };
                    if (mm.MaxRow > 0 && mm.MaxColumn > 0)
                    {
                        for (Int32 ro = mm.MinRow; ro <= mm.MaxRow; ro++)
                        {
                            for (Int32 co = mm.MinColumn; co <= mm.MaxColumn; co++)
                            {
                                var address = new XLAddress(range.Worksheet, ro, co, false, false);
                                if (!hash.Contains(address))
                                    hash.Add(address);
                            }
                        }
                    }
                }
            }

            if (_usedCellsOnly)
            {
                foreach (KeyValuePair<XLWorksheet, HashSet<IXLAddress>> cir in cellsInRanges)
                {
                    var cellsCollection = cir.Key.Internals.CellsCollection;
                    foreach (IXLAddress a in cir.Value)
                    {
                        if (cellsCollection.ContainsKey(a))
                        {
                            var cell = cellsCollection[a];
                            if (!StringExtensions.IsNullOrWhiteSpace((cell).InnerText)
                                || (_includeFormats && (!cell.Style.Equals(cir.Key.Style) || cell.IsMerged())))
                                yield return cell;
                        }
                    }
                }
            }
            else
            {
                foreach (KeyValuePair<XLWorksheet, HashSet<IXLAddress>> cir in cellsInRanges)
                {
                    foreach (IXLAddress address in cir.Value)
                        yield return cir.Key.Cell(address);
                }
            }
        }

        #endregion

        #region IXLCells Members

        IEnumerator<IXLCell> IEnumerable<IXLCell>.GetEnumerator()
        {
            foreach (XLCell cell in this)
                yield return cell;
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public IXLStyle Style
        {
            get { return _style; }
            set
            {
                _style = new XLStyle(this, value);
                this.ForEach<XLCell>(c => c.Style = _style);
            }
        }

        public Object Value
        {
            set { this.ForEach<XLCell>(c => c.Value = value); }
        }

        public IXLCells SetDataType(XLCellValues dataType)
        {
            this.ForEach<XLCell>(c => c.DataType = dataType);
            return this;
        }

        public XLCellValues DataType
        {
            set { this.ForEach<XLCell>(c => c.DataType = value); }
        }

        public void Clear()
        {
            this.ForEach<XLCell>(c => c.Clear());
        }

        public void ClearStyles()
        {
            this.ForEach<XLCell>(c => c.ClearStyles());
        }

        public String FormulaA1
        {
            set { this.ForEach<XLCell>(c => c.FormulaA1 = value); }
        }

        public String FormulaR1C1
        {
            set { this.ForEach<XLCell>(c => c.FormulaR1C1 = value); }
        }

        #endregion

        #region IXLStylized Members

        public IEnumerable<IXLStyle> Styles
        {
            get
            {
                UpdatingStyle = true;
                yield return _style;
                foreach (XLCell c in this)
                    yield return c.Style;
                UpdatingStyle = false;
            }
        }

        public Boolean UpdatingStyle { get; set; }

        public IXLStyle InnerStyle
        {
            get { return _style; }
            set { _style = new XLStyle(this, value); }
        }

        public IXLRanges RangesUsed
        {
            get
            {
                var retVal = new XLRanges();
                this.ForEach<XLCell>(c => retVal.Add(c.AsRange()));
                return retVal;
            }
        }

        #endregion

        public void Add(XLRangeAddress rangeAddress)
        {
            _rangeAddresses.Add(rangeAddress);
        }

        public void Add(XLCell cell)
        {
            _rangeAddresses.Add(new XLRangeAddress(cell.Address, cell.Address));
        }

        //--

        #region Nested type: MinMax

        private struct MinMax
        {
            public Int32 MaxColumn;
            public Int32 MaxRow;
            public Int32 MinColumn;
            public Int32 MinRow;
        }

        #endregion
    }
}