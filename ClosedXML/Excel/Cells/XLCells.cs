using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    using System.Linq;

    internal class XLCells : IXLCells, IXLStylized, IEnumerable<XLCell>
    {
        public Boolean StyleChanged { get; set; }
        #region Fields

        private readonly bool _includeFormats;
        private readonly List<XLRangeAddress> _rangeAddresses = new List<XLRangeAddress>();
        private readonly bool _usedCellsOnly;
        private IXLStyle _style;
        private readonly Func<IXLCell, Boolean> _predicate;
        #endregion

        #region Constructor

        public XLCells(bool usedCellsOnly, bool includeFormats, Func<IXLCell, Boolean> predicate = null)
        {
            _style = new XLStyle(this, XLWorkbook.DefaultStyle);
            _usedCellsOnly = usedCellsOnly;
            _includeFormats = includeFormats;
            _predicate = predicate;
        }

        #endregion

        #region IEnumerable<XLCell> Members

        public IEnumerator<XLCell> GetEnumerator()
        {
            var cellsInRanges = new Dictionary<XLWorksheet, HashSet<XLSheetPoint>>();
            Boolean oneRange = _rangeAddresses.Count == 1;
            foreach (XLRangeAddress range in _rangeAddresses)
            {
                HashSet<XLSheetPoint> hash;
                if (cellsInRanges.ContainsKey(range.Worksheet))
                    hash = cellsInRanges[range.Worksheet];
                else
                {
                    hash = new HashSet<XLSheetPoint>();
                    cellsInRanges.Add(range.Worksheet, hash);
                }

                if (_usedCellsOnly)
                {
                    if (oneRange)
                    {
                        var cellRange = range.Worksheet.Internals.CellsCollection
                                                .GetCells(
                                                range.FirstAddress.RowNumber,
                                                range.FirstAddress.ColumnNumber,
                                                range.LastAddress.RowNumber,
                                                range.LastAddress.ColumnNumber)
                                                .Where(c => 
                                                            !c.IsEmpty(_includeFormats) 
                                                            && (_predicate == null || _predicate(c))
                                                            );

                        foreach(var cell in cellRange)
                        {
                            yield return cell;
                        }
                    }
                    else
                    {
                        var tmpRange = range;
                        var addressList = range.Worksheet.Internals.CellsCollection
                            .GetSheetPoints(
                            tmpRange.FirstAddress.RowNumber,
                            tmpRange.FirstAddress.ColumnNumber,
                            tmpRange.LastAddress.RowNumber,
                            tmpRange.LastAddress.ColumnNumber);

                        foreach (XLSheetPoint a in addressList.Where(a => !hash.Contains(a)))
                        {
                            hash.Add(a);
                        }
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
                                if (oneRange)
                                {
                                    var c = range.Worksheet.Cell(ro, co);
                                    if (_predicate == null || _predicate(c))
                                        yield return c;
                                }
                                else
                                {
                                    var address = new XLSheetPoint(ro, co);
                                    if (!hash.Contains(address))
                                        hash.Add(address);
                                }
                            }
                        }
                    }
                }
            }

            if (!oneRange)
            {
                if (_usedCellsOnly)
                {
                    var cellRange = cellsInRanges.SelectMany(
                                cir =>
                                cir.Value.Select(a => cir.Key.Internals.CellsCollection.GetCell(a)).Where(
                                    cell => cell != null && (
                                                                !cell.IsEmpty(_includeFormats) 
                                                                && (_predicate == null || _predicate(cell))
                                                                )));

                    foreach (var cell in cellRange)
                    {
                        yield return cell;
                    }
                }
                else
                {
                    foreach (var cir in cellsInRanges)
                    {
                        foreach (XLSheetPoint a in cir.Value)
                        {
                            var c = cir.Key.Cell(a.Row, a.Column);
                            if (_predicate == null || _predicate(c))
                                yield return c;
                        }
                    }
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


        public IXLCells Clear(XLClearOptions clearOptions = XLClearOptions.ContentsAndFormats)
        {
            this.ForEach<XLCell>(c => c.Clear(clearOptions));
            return this;
        }

        public void DeleteComments() {
            this.ForEach<XLCell>(c => c.DeleteComment());
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

        public void Select()
        {
            foreach (var cell in this)
                cell.Select();
        }
    }
}