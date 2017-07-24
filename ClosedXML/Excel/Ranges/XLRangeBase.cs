using ClosedXML.Excel.Misc;
using ClosedXML.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;



namespace ClosedXML.Excel
{
    internal abstract class XLRangeBase : IXLRangeBase, IXLStylized
    {
        public Boolean StyleChanged { get; set; }
        #region Fields

        private IXLStyle _style;
        private XLSortElements _sortRows;
        private XLSortElements _sortColumns;

        #endregion

        private Int32 _styleCacheId;
        protected void SetStyle(IXLStyle styleToUse)
        {
            _styleCacheId = Worksheet.Workbook.GetStyleId(styleToUse);
            _style = null;
            StyleChanged = false;
        }
        protected void SetStyle(Int32 styleId)
        {
            _styleCacheId = styleId;
            _style = null;
            StyleChanged = false;
        }
        public Int32 GetStyleId()
        {
            if (StyleChanged)
                SetStyle(Style);

            return _styleCacheId;
        }
        protected IXLStyle GetStyle()
        {
            return _style ?? (_style = new XLStyle(this, Worksheet.Workbook.GetStyleById(_styleCacheId)));
        }

        #region Constructor

        static Int32 IdCounter = 0;
        readonly Int32 Id;

        protected XLRangeBase(XLRangeAddress rangeAddress)
        {

            Id = ++IdCounter;

            RangeAddress = new XLRangeAddress(rangeAddress);
        }

        #endregion

        private XLCallbackAction _shiftedRowsAction;

        protected void SubscribeToShiftedRows(Action<XLRange, Int32> action)
        {
            if (Worksheet == null || !Worksheet.EventTrackingEnabled) return;

            _shiftedRowsAction = new XLCallbackAction(action);

            RangeAddress.Worksheet.RangeShiftedRows.Add(_shiftedRowsAction);
        }

        private XLCallbackAction _shiftedColumnsAction;
        protected void SubscribeToShiftedColumns(Action<XLRange, Int32> action)
        {
            if (Worksheet == null || !Worksheet.EventTrackingEnabled) return;

            _shiftedColumnsAction = new XLCallbackAction(action);

            RangeAddress.Worksheet.RangeShiftedColumns.Add(_shiftedColumnsAction);
        }

        #region Public properties

        //public XLRangeAddress RangeAddress { get; protected set; }

        private XLRangeAddress _rangeAddress;
        public XLRangeAddress RangeAddress
        {
            get { return _rangeAddress; }
            protected set { _rangeAddress = value; }
        }

        public XLWorksheet Worksheet
        {
            get { return RangeAddress.Worksheet; }
        }

        public XLDataValidation NewDataValidation
        {
            get
            {
                var newRanges = new XLRanges { AsRange() };
                var dataValidation = new XLDataValidation(newRanges);

                Worksheet.DataValidations.Add(dataValidation);
                return dataValidation;
            }
        }

        public XLDataValidation DataValidation
        {
            get
            {
                IXLDataValidation dataValidationToCopy = null;
                var dvEmpty = new List<IXLDataValidation>();
                foreach (IXLDataValidation dv in Worksheet.DataValidations)
                {
                    foreach (IXLRange dvRange in dv.Ranges.Where(dvRange => dvRange.Intersects(this)))
                    {
                        if (dataValidationToCopy == null)
                            dataValidationToCopy = dv;

                        dv.Ranges.Remove(dvRange);
                        foreach (var column in dvRange.Columns())
                        {
                            if (column.Intersects(this))
                            {
                                Int32 dvStart = column.RangeAddress.FirstAddress.RowNumber;
                                Int32 dvEnd = column.RangeAddress.LastAddress.RowNumber;
                                Int32 thisStart = RangeAddress.FirstAddress.RowNumber;
                                Int32 thisEnd = RangeAddress.LastAddress.RowNumber;

                                if (thisStart > dvStart && thisEnd < dvEnd)
                                {
                                    var r1 = Worksheet.Column(column.ColumnNumber()).Column(dvStart, thisStart - 1);
                                    r1.Dispose();
                                    dv.Ranges.Add(r1);
                                    var r2 = Worksheet.Column(column.ColumnNumber()).Column(thisEnd + 1, dvEnd);
                                    r2.Dispose();
                                    dv.Ranges.Add(r2);
                                }
                                else
                                {
                                    Int32 coStart;
                                    if (dvStart < thisStart)
                                        coStart = dvStart;
                                    else
                                        coStart = thisEnd + 1;

                                    if (coStart <= dvEnd)
                                    {
                                        Int32 coEnd;
                                        if (dvEnd > thisEnd)
                                            coEnd = dvEnd;
                                        else
                                            coEnd = thisStart - 1;

                                        if (coEnd >= dvStart)
                                        {
                                            var r = Worksheet.Column(column.ColumnNumber()).Column(coStart, coEnd);
                                            r.Dispose();
                                            dv.Ranges.Add(r);
                                        }
                                    }
                                }
                            }
                            else
                            {
                                column.Dispose();
                                dv.Ranges.Add(column);
                            }
                        }

                        if (!dv.Ranges.Any())
                            dvEmpty.Add(dv);
                    }
                }

                dvEmpty.ForEach(dv => Worksheet.DataValidations.Delete(dv));

                var newRanges = new XLRanges { AsRange() };
                var dataValidation = new XLDataValidation(newRanges);
                if (dataValidationToCopy != null)
                    dataValidation.CopyFrom(dataValidationToCopy);

                Worksheet.DataValidations.Add(dataValidation);
                return dataValidation;
            }
        }

        #region IXLRangeBase Members

        IXLRangeAddress IXLRangeBase.RangeAddress
        {
            get { return RangeAddress; }
        }

        IXLWorksheet IXLRangeBase.Worksheet
        {
            get { return RangeAddress.Worksheet; }
        }

        public String FormulaA1
        {
            set
            {
                Cells().ForEach(c =>
                                    {
                                        c.FormulaA1 = value;
                                        c.FormulaReference = RangeAddress;
                                    });
            }
        }

        public String FormulaR1C1
        {
            set
            {
                Cells().ForEach(c =>
                {
                    c.FormulaR1C1 = value;
                    c.FormulaReference = RangeAddress;
                });
            }
        }

        public Boolean ShareString
        {
            set { Cells().ForEach(c => c.ShareString = value); }
        }

        public IXLHyperlinks Hyperlinks
        {
            get
            {
                var hyperlinks = new XLHyperlinks();
                var hls = from hl in Worksheet.Hyperlinks
                          where Contains(hl.Cell.AsRange())
                          select hl;
                hls.ForEach(hyperlinks.Add);
                return hyperlinks;
            }
        }



        public Object Value
        {
            set { Cells().ForEach(c => c.Value = value); }
        }

        public XLCellValues DataType
        {
            set { Cells().ForEach(c => c.DataType = value); }
        }

        #endregion

        #region IXLStylized Members

        public IXLRanges RangesUsed
        {
            get
            {
                var retVal = new XLRanges { AsRange() };
                return retVal;
            }
        }

        #endregion

        #endregion

        #region IXLRangeBase Members

        IXLCell IXLRangeBase.FirstCell()
        {
            return FirstCell();
        }

        IXLCell IXLRangeBase.LastCell()
        {
            return LastCell();
        }

        IXLCell IXLRangeBase.FirstCellUsed()
        {
            return FirstCellUsed(false);
        }

        IXLCell IXLRangeBase.FirstCellUsed(bool includeFormats)
        {
            return FirstCellUsed(includeFormats);
        }

        IXLCell IXLRangeBase.LastCellUsed()
        {
            return LastCellUsed(false);
        }

        IXLCell IXLRangeBase.LastCellUsed(bool includeFormats)
        {
            return LastCellUsed(includeFormats);
        }

        public IXLCells Cells()
        {
            return Cells(false);
        }

        public IXLCells Cells(Boolean usedCellsOnly)
        {
            return Cells(usedCellsOnly, false);
        }

        public IXLCells Cells(Boolean usedCellsOnly, Boolean includeFormats)
        {
            var cells = new XLCells(usedCellsOnly, includeFormats) { RangeAddress };
            return cells;
        }

        public IXLCells Cells(String cells)
        {
            return Ranges(cells).Cells();
        }

        public IXLCells Cells(Func<IXLCell, Boolean> predicate)
        {
            var cells = new XLCells(false, false, predicate) { RangeAddress };
            return cells;
        }

        public IXLCells CellsUsed()
        {
            return Cells(true);
        }

        public IXLRange Merge()
        {
            return Merge(true);
        }

        public IXLRange Merge(Boolean checkIntersect)
        {
            if (checkIntersect)
            {
                using (IXLRange range = Worksheet.Range(RangeAddress))
                {
                    foreach (var mergedRange in Worksheet.Internals.MergedRanges)
                    {
                        if (mergedRange.Intersects(range))
                        {
                            Worksheet.Internals.MergedRanges.Remove(mergedRange);
                        }
                    }
                }
            }

            var asRange = AsRange();
            Worksheet.Internals.MergedRanges.Add(asRange);

            return asRange;
        }

        public IXLRange Unmerge()
        {
            string tAddress = RangeAddress.ToString();
            var asRange = AsRange();
            if (Worksheet.Internals.MergedRanges.Select(m => m.RangeAddress.ToString()).Any(mAddress => mAddress == tAddress))
                Worksheet.Internals.MergedRanges.Remove(asRange);

            return asRange;
        }

        public IXLRangeBase Clear(XLClearOptions clearOptions = XLClearOptions.ContentsAndFormats)
        {
            var includeFormats = clearOptions == XLClearOptions.Formats ||
                                 clearOptions == XLClearOptions.ContentsAndFormats;
            foreach (var cell in CellsUsed(includeFormats))
            {
                (cell as XLCell).Clear(clearOptions, true);
            }

            if (includeFormats)
            {
                ClearMerged();
            }

            if (clearOptions == XLClearOptions.ContentsAndFormats)
            {
                Worksheet.Internals.CellsCollection.RemoveAll(
                    RangeAddress.FirstAddress.RowNumber,
                    RangeAddress.FirstAddress.ColumnNumber,
                    RangeAddress.LastAddress.RowNumber,
                    RangeAddress.LastAddress.ColumnNumber
                    );
            }
            return this;
        }

        public void DeleteComments()
        {
            Cells().DeleteComments();
        }

        public bool Contains(String rangeAddress)
        {
            string addressToUse = rangeAddress.Contains("!")
                                      ? rangeAddress.Substring(rangeAddress.IndexOf("!") + 1)
                                      : rangeAddress;

            XLAddress firstAddress;
            XLAddress lastAddress;
            if (addressToUse.Contains(':'))
            {
                var arrRange = addressToUse.Split(':');
                firstAddress = XLAddress.Create(Worksheet, arrRange[0]);
                lastAddress = XLAddress.Create(Worksheet, arrRange[1]);
            }
            else
            {
                firstAddress = XLAddress.Create(Worksheet, addressToUse);
                lastAddress = XLAddress.Create(Worksheet, addressToUse);
            }
            return Contains(firstAddress, lastAddress);
        }

        public bool Contains(IXLRangeBase range)
        {
            return Contains((XLAddress)range.RangeAddress.FirstAddress, (XLAddress)range.RangeAddress.LastAddress);
        }

        public bool Intersects(string rangeAddress)
        {
            using (var range = Worksheet.Range(rangeAddress))
                return Intersects(range);
        }

        public bool Intersects(IXLRangeBase range)
        {
            if (range.RangeAddress.IsInvalid || RangeAddress.IsInvalid)
                return false;
            var ma = range.RangeAddress;
            var ra = RangeAddress;

            return !( // See if the two ranges intersect...
                    ma.FirstAddress.ColumnNumber > ra.LastAddress.ColumnNumber
                    || ma.LastAddress.ColumnNumber < ra.FirstAddress.ColumnNumber
                    || ma.FirstAddress.RowNumber > ra.LastAddress.RowNumber
                    || ma.LastAddress.RowNumber < ra.FirstAddress.RowNumber
                    );
        }

        public virtual IXLStyle Style
        {
            get { return GetStyle(); }
            set { Cells().ForEach(c => c.Style = value); }
        }
        IXLRange IXLRangeBase.AsRange()
        {
            return AsRange();
        }
        public virtual XLRange AsRange()
        {
            return Worksheet.Range(RangeAddress.FirstAddress, RangeAddress.LastAddress);
        }



        public IXLRange AddToNamed(String rangeName)
        {
            return AddToNamed(rangeName, XLScope.Workbook);
        }

        public IXLRange AddToNamed(String rangeName, XLScope scope)
        {
            return AddToNamed(rangeName, scope, null);
        }

        public IXLRange AddToNamed(String rangeName, XLScope scope, String comment)
        {
            var namedRanges = scope == XLScope.Workbook
                                  ? Worksheet.Workbook.NamedRanges
                                  : Worksheet.NamedRanges;

            if (namedRanges.Any(nr => String.Compare(nr.Name, rangeName, true) == 0))
            {
                var namedRange = namedRanges.Single(nr => String.Compare(nr.Name, rangeName, true) == 0);
                namedRange.Add(Worksheet.Workbook, RangeAddress.ToStringFixed(XLReferenceStyle.A1, true));
            }
            else
                namedRanges.Add(rangeName, RangeAddress.ToStringFixed(XLReferenceStyle.A1, true), comment);
            return AsRange();
        }

        public IXLRangeBase SetValue<T>(T value)
        {
            Cells().ForEach(c => c.SetValue(value));
            return this;
        }

        public Boolean IsMerged()
        {
            return Cells().Any(c => c.IsMerged());
        }

        public Boolean IsEmpty()
        {
            return !CellsUsed().Any() || CellsUsed().Any(c => c.IsEmpty());
        }

        public Boolean IsEmpty(Boolean includeFormats)
        {
            return !CellsUsed(includeFormats).Cast<XLCell>().Any() ||
                   CellsUsed(includeFormats).Cast<XLCell>().Any(c => c.IsEmpty(includeFormats));
        }

        #endregion

        #region IXLStylized Members

        public virtual IEnumerable<IXLStyle> Styles
        {
            get
            {
                UpdatingStyle = true;
                foreach (IXLCell cell in Cells())
                    yield return cell.Style;
                UpdatingStyle = false;
            }
        }

        public virtual Boolean UpdatingStyle { get; set; }

        public virtual IXLStyle InnerStyle
        {
            get { return GetStyle(); }
            set { SetStyle(value); }
        }

        #endregion

        public XLCell FirstCell()
        {
            return Cell(1, 1);
        }

        public XLCell LastCell()
        {
            return Cell(RowCount(), ColumnCount());
        }

        public XLCell FirstCellUsed()
        {
            return FirstCellUsed(false, null);
        }

        public XLCell FirstCellUsed(Boolean includeFormats)
        {
            return FirstCellUsed(includeFormats, null);
        }

        IXLCell IXLRangeBase.FirstCellUsed(Func<IXLCell, Boolean> predicate)
        {
            return FirstCellUsed(predicate);
        }

        public XLCell FirstCellUsed(Func<IXLCell, Boolean> predicate)
        {
            return FirstCellUsed(false, predicate);
        }

        IXLCell IXLRangeBase.FirstCellUsed(Boolean includeFormats, Func<IXLCell, Boolean> predicate)
        {
            return FirstCellUsed(includeFormats, predicate);
        }

        public XLCell FirstCellUsed(Boolean includeFormats, Func<IXLCell, Boolean> predicate)
        {
            Int32 fRow = RangeAddress.FirstAddress.RowNumber;
            Int32 lRow = RangeAddress.LastAddress.RowNumber;
            Int32 fColumn = RangeAddress.FirstAddress.ColumnNumber;
            Int32 lColumn = RangeAddress.LastAddress.ColumnNumber;

            var sp = Worksheet.Internals.CellsCollection.FirstPointUsed(fRow, fColumn, lRow, lColumn, includeFormats, predicate);

            if (includeFormats)
            {
                var rowsUsed =
                    Worksheet.Internals.RowsCollection.Where(r => r.Key >= fRow && r.Key <= lRow && !r.Value.IsEmpty(true));

                var columnsUsed =
                    Worksheet.Internals.ColumnsCollection.Where(c => c.Key >= fColumn && c.Key <= lColumn && !c.Value.IsEmpty(true));

                // If there's a row or a column then check if the style is different
                // and pick the first cell and check the style of it, if different
                // than default then it's your cell.

                Int32 ro = 0;
                if (rowsUsed.Any())
                    if (sp.Row > 0)
                        ro = Math.Min(sp.Row, rowsUsed.First().Key);
                    else
                        ro = rowsUsed.First().Key;

                Int32 co = 0;
                if (columnsUsed.Any())
                    if (sp.Column > 0)
                        co = Math.Min(sp.Column, columnsUsed.First().Key);
                    else
                        co = columnsUsed.First().Key;

                if (ro > 0 && co > 0)
                    return Worksheet.Cell(ro, co);

                if (ro > 0 && lColumn < XLHelper.MaxColumnNumber)
                {
                    for (co = fColumn; co <= lColumn; co++)
                    {
                        var cell = Worksheet.Cell(ro, co);
                        if (!cell.IsEmpty(true)) return cell;
                    }
                }
                else if (co > 0 && lRow < XLHelper.MaxRowNumber)
                {
                    for (ro = fRow; ro <= lRow; ro++)
                    {
                        var cell = Worksheet.Cell(ro, co);
                        if (!cell.IsEmpty(true)) return cell;
                    }
                }

                if (Worksheet.MergedRanges.Any(r => r.Intersects(this)))
                {
                    Int32 minRo =
                        Worksheet.MergedRanges.Where(r => r.Intersects(this)).Min(r => r.RangeAddress.FirstAddress.RowNumber);
                    Int32 minCo =
                        Worksheet.MergedRanges.Where(r => r.Intersects(this)).Min(r => r.RangeAddress.FirstAddress.ColumnNumber);

                    return Worksheet.Cell(minRo, minCo);
                }
            }


            if (sp.Row > 0)
                return Worksheet.Cell(sp.Row, sp.Column);

            return null;
        }

        public XLCell LastCellUsed()
        {
            return LastCellUsed(false, null);
        }

        public XLCell LastCellUsed(Boolean includeFormats)
        {
            return LastCellUsed(includeFormats, null);
        }

        IXLCell IXLRangeBase.LastCellUsed(Func<IXLCell, Boolean> predicate)
        {
            return LastCellUsed(predicate);
        }

        public XLCell LastCellUsed(Func<IXLCell, Boolean> predicate)
        {
            return LastCellUsed(false, predicate);
        }

        IXLCell IXLRangeBase.LastCellUsed(Boolean includeFormats, Func<IXLCell, Boolean> predicate)
        {
            return LastCellUsed(includeFormats, predicate);
        }

        public XLCell LastCellUsed(Boolean includeFormats, Func<IXLCell, Boolean> predicate)
        {
            Int32 fRow = RangeAddress.FirstAddress.RowNumber;
            Int32 lRow = RangeAddress.LastAddress.RowNumber;
            Int32 fColumn = RangeAddress.FirstAddress.ColumnNumber;
            Int32 lColumn = RangeAddress.LastAddress.ColumnNumber;

            var sp = Worksheet.Internals.CellsCollection.LastPointUsed(fRow, fColumn, lRow, lColumn, includeFormats, predicate);

            if (includeFormats)
            {
                var rowsUsed =
                    Worksheet.Internals.RowsCollection.Where(r => r.Key >= fRow && r.Key <= lRow && !r.Value.IsEmpty(true));

                var columnsUsed =
                    Worksheet.Internals.ColumnsCollection.Where(c => c.Key >= fColumn && c.Key <= lColumn && !c.Value.IsEmpty(true));

                // If there's a row or a column then check if the style is different
                // and pick the first cell and check the style of it, if different
                // than default then it's your cell.

                Int32 ro = 0;
                if (rowsUsed.Any())
                    ro = Math.Max(sp.Row, rowsUsed.Last().Key);

                Int32 co = 0;
                if (columnsUsed.Any())
                    co = Math.Max(sp.Column, columnsUsed.Last().Key);

                if (ro > 0 && co > 0)
                    return Worksheet.Cell(ro, co);

                if (ro > 0 && lColumn < XLHelper.MaxColumnNumber)
                {
                    for (co = lColumn; co >= fColumn; co--)
                    {
                        var cell = Worksheet.Cell(ro, co);
                        if (!cell.IsEmpty(true)) return cell;
                    }
                }
                else if (co > 0 && lRow < XLHelper.MaxRowNumber)
                {
                    for (ro = lRow; ro >= fRow; ro--)
                    {
                        var cell = Worksheet.Cell(ro, co);
                        if (!cell.IsEmpty(true)) return cell;
                    }
                }

                if (Worksheet.MergedRanges.Any(r => r.Intersects(this)))
                {
                    Int32 minRo =
                        Worksheet.MergedRanges.Where(r => r.Intersects(this)).Max(r => r.RangeAddress.LastAddress.RowNumber);
                    Int32 minCo =
                        Worksheet.MergedRanges.Where(r => r.Intersects(this)).Max(r => r.RangeAddress.LastAddress.ColumnNumber);

                    return Worksheet.Cell(minRo, minCo);
                }
            }


            if (sp.Row > 0)
                return Worksheet.Cell(sp.Row, sp.Column);

            return null;
        }

        public XLCell Cell(Int32 row, Int32 column)
        {
            return Cell(new XLAddress(Worksheet, row, column, false, false));
        }

        public XLCell Cell(String cellAddressInRange)
        {

            if (XLHelper.IsValidA1Address(cellAddressInRange))
                return Cell(XLAddress.Create(Worksheet, cellAddressInRange));

            return (XLCell)Worksheet.NamedRange(cellAddressInRange).Ranges.First().FirstCell();
        }

        public XLCell Cell(Int32 row, String column)
        {
            return Cell(new XLAddress(Worksheet, row, column, false, false));
        }

        public XLCell Cell(IXLAddress cellAddressInRange)
        {
            return Cell(cellAddressInRange.RowNumber, cellAddressInRange.ColumnNumber);
        }

        public XLCell Cell(XLAddress cellAddressInRange)
        {
            Int32 absRow = cellAddressInRange.RowNumber + RangeAddress.FirstAddress.RowNumber - 1;
            Int32 absColumn = cellAddressInRange.ColumnNumber + RangeAddress.FirstAddress.ColumnNumber - 1;

            if (absRow <= 0 || absRow > XLHelper.MaxRowNumber)
            {
                throw new IndexOutOfRangeException(String.Format("Row number must be between 1 and {0}",
                                                                 XLHelper.MaxRowNumber));
            }

            if (absColumn <= 0 || absColumn > XLHelper.MaxColumnNumber)
            {
                throw new IndexOutOfRangeException(String.Format("Column number must be between 1 and {0}",
                                                                 XLHelper.MaxColumnNumber));
            }

            var cell = Worksheet.Internals.CellsCollection.GetCell(absRow,
                                                                   absColumn);

            if (cell != null)
                return cell;

            Int32 styleId = GetStyleId();
            Int32 worksheetStyleId = Worksheet.GetStyleId();

            if (styleId == worksheetStyleId)
            {
                XLRow row;
                XLColumn column;
                if (Worksheet.Internals.RowsCollection.TryGetValue(absRow, out row)
                    && row.GetStyleId() != worksheetStyleId)
                    styleId = row.GetStyleId();
                else if (Worksheet.Internals.ColumnsCollection.TryGetValue(absColumn, out column)
                    && column.GetStyleId() != worksheetStyleId)
                    styleId = column.GetStyleId();
            }
            var absoluteAddress = new XLAddress(this.Worksheet,
                                 absRow,
                                 absColumn,
                                 cellAddressInRange.FixedRow,
                                 cellAddressInRange.FixedColumn);

            Int32 newCellStyleId = styleId;

            // If the default style for this range base is empty, but the worksheet 
            // has a default style, use the worksheet's default style
            if (styleId == 0 && worksheetStyleId != 0)
                newCellStyleId = worksheetStyleId;

            var newCell = new XLCell(Worksheet, absoluteAddress, newCellStyleId);
            Worksheet.Internals.CellsCollection.Add(absRow, absColumn, newCell);
            return newCell;
        }

        public Int32 RowCount()
        {
            return RangeAddress.LastAddress.RowNumber - RangeAddress.FirstAddress.RowNumber + 1;
        }

        public Int32 RowNumber()
        {
            return RangeAddress.FirstAddress.RowNumber;
        }

        public Int32 ColumnCount()
        {
            return RangeAddress.LastAddress.ColumnNumber - RangeAddress.FirstAddress.ColumnNumber + 1;
        }

        public Int32 ColumnNumber()
        {
            return RangeAddress.FirstAddress.ColumnNumber;
        }

        public String ColumnLetter()
        {
            return RangeAddress.FirstAddress.ColumnLetter;
        }

        public virtual XLRange Range(String rangeAddressStr)
        {
            var rangeAddress = new XLRangeAddress(Worksheet, rangeAddressStr);
            return Range(rangeAddress);
        }

        public XLRange Range(IXLCell firstCell, IXLCell lastCell)
        {
            var newFirstCellAddress = firstCell.Address as XLAddress;
            var newLastCellAddress = lastCell.Address as XLAddress;

            return GetRange(newFirstCellAddress, newLastCellAddress);
        }

        private XLRange GetRange(XLAddress newFirstCellAddress, XLAddress newLastCellAddress)
        {
            var newRangeAddress = new XLRangeAddress(newFirstCellAddress, newLastCellAddress);
            var xlRangeParameters = new XLRangeParameters(newRangeAddress, Style);
            if (
                newFirstCellAddress.RowNumber < RangeAddress.FirstAddress.RowNumber
                || newFirstCellAddress.RowNumber > RangeAddress.LastAddress.RowNumber
                || newLastCellAddress.RowNumber > RangeAddress.LastAddress.RowNumber
                || newFirstCellAddress.ColumnNumber < RangeAddress.FirstAddress.ColumnNumber
                || newFirstCellAddress.ColumnNumber > RangeAddress.LastAddress.ColumnNumber
                || newLastCellAddress.ColumnNumber > RangeAddress.LastAddress.ColumnNumber
                )
            {
                throw new ArgumentOutOfRangeException(String.Format(
                    "The cells {0} and {1} are outside the range '{2}'.",
                    newFirstCellAddress,
                    newLastCellAddress,
                    ToString()));
            }

            return new XLRange(xlRangeParameters);
        }

        public XLRange Range(String firstCellAddress, String lastCellAddress)
        {
            var rangeAddress = new XLRangeAddress(XLAddress.Create(Worksheet, firstCellAddress),
                                                  XLAddress.Create(Worksheet, lastCellAddress));
            return Range(rangeAddress);
        }

        public XLRange Range(Int32 firstCellRow, Int32 firstCellColumn, Int32 lastCellRow, Int32 lastCellColumn)
        {
            var rangeAddress = new XLRangeAddress(new XLAddress(Worksheet, firstCellRow, firstCellColumn, false, false),
                                                  new XLAddress(Worksheet, lastCellRow, lastCellColumn, false, false));
            return Range(rangeAddress);
        }

        public XLRange Range(IXLAddress firstCellAddress, IXLAddress lastCellAddress)
        {
            var rangeAddress = new XLRangeAddress(firstCellAddress as XLAddress, lastCellAddress as XLAddress);
            return Range(rangeAddress);
        }

        public XLRange Range(IXLRangeAddress rangeAddress)
        {

            var newFirstCellAddress = new XLAddress((XLWorksheet)rangeAddress.FirstAddress.Worksheet,
                                 rangeAddress.FirstAddress.RowNumber + RangeAddress.FirstAddress.RowNumber - 1,
                                 rangeAddress.FirstAddress.ColumnNumber + RangeAddress.FirstAddress.ColumnNumber - 1,
                                 rangeAddress.FirstAddress.FixedRow,
                                 rangeAddress.FirstAddress.FixedColumn);

            newFirstCellAddress.FixedRow = rangeAddress.FirstAddress.FixedRow;
            newFirstCellAddress.FixedColumn = rangeAddress.FirstAddress.FixedColumn;

            var newLastCellAddress = new XLAddress((XLWorksheet)rangeAddress.LastAddress.Worksheet,
                                rangeAddress.LastAddress.RowNumber + RangeAddress.FirstAddress.RowNumber - 1,
                                rangeAddress.LastAddress.ColumnNumber + RangeAddress.FirstAddress.ColumnNumber - 1,
                                rangeAddress.LastAddress.FixedRow,
                                rangeAddress.LastAddress.FixedColumn);

            newLastCellAddress.FixedRow = rangeAddress.LastAddress.FixedRow;
            newLastCellAddress.FixedColumn = rangeAddress.LastAddress.FixedColumn;

            return GetRange(newFirstCellAddress, newLastCellAddress);
        }

        public IXLRanges Ranges(String ranges)
        {
            var retVal = new XLRanges();
            var rangePairs = ranges.Split(',');
            foreach (string pair in rangePairs)
                retVal.Add(Range(pair.Trim()));
            return retVal;
        }

        public IXLRanges Ranges(params String[] ranges)
        {
            var retVal = new XLRanges();
            foreach (string pair in ranges)
                retVal.Add(Range(pair));
            return retVal;
        }

        protected String FixColumnAddress(String address)
        {
            Int32 test;
            if (Int32.TryParse(address, out test))
                return "A" + address;
            return address;
        }

        protected String FixRowAddress(String address)
        {
            Int32 test;
            if (Int32.TryParse(address, out test))
                return XLHelper.GetColumnLetterFromNumber(test) + "1";
            return address;
        }

        public IXLCells CellsUsed(bool includeFormats)
        {
            var cells = new XLCells(true, includeFormats) { RangeAddress };
            return cells;
        }

        public IXLCells CellsUsed(Func<IXLCell, Boolean> predicate)
        {
            var cells = new XLCells(true, false, predicate) { RangeAddress };
            return cells;
        }

        public IXLCells CellsUsed(Boolean includeFormats, Func<IXLCell, Boolean> predicate)
        {
            var cells = new XLCells(true, includeFormats, predicate) { RangeAddress };
            return cells;
        }

        public IXLRangeColumns InsertColumnsAfter(Int32 numberOfColumns)
        {
            return InsertColumnsAfter(numberOfColumns, true);
        }

        public IXLRangeColumns InsertColumnsAfter(Int32 numberOfColumns, Boolean expandRange)
        {
            var retVal = InsertColumnsAfter(false, numberOfColumns);
            // Adjust the range
            if (expandRange)
            {
                RangeAddress = new XLRangeAddress(
                    new XLAddress(Worksheet,
                                  RangeAddress.FirstAddress.RowNumber,
                                  RangeAddress.FirstAddress.ColumnNumber,
                                  RangeAddress.FirstAddress.FixedRow,
                                  RangeAddress.FirstAddress.FixedColumn),
                    new XLAddress(Worksheet,
                                  RangeAddress.LastAddress.RowNumber,
                                  RangeAddress.LastAddress.ColumnNumber + numberOfColumns,
                                  RangeAddress.LastAddress.FixedRow,
                                  RangeAddress.LastAddress.FixedColumn));
            }
            return retVal;
        }

        public IXLRangeColumns InsertColumnsAfter(Boolean onlyUsedCells, Int32 numberOfColumns, Boolean formatFromLeft = true)
        {
            return InsertColumnsAfterInternal(onlyUsedCells, numberOfColumns, formatFromLeft);
        }

        public void InsertColumnsAfterVoid(Boolean onlyUsedCells, Int32 numberOfColumns, Boolean formatFromLeft = true)
        {
            InsertColumnsAfterInternal(onlyUsedCells, numberOfColumns, formatFromLeft, nullReturn: true);
        }

        private IXLRangeColumns InsertColumnsAfterInternal(Boolean onlyUsedCells, Int32 numberOfColumns, Boolean formatFromLeft = true, Boolean nullReturn = false)
        {
            int columnCount = ColumnCount();
            int firstColumn = RangeAddress.FirstAddress.ColumnNumber + columnCount;
            if (firstColumn > XLHelper.MaxColumnNumber)
                firstColumn = XLHelper.MaxColumnNumber;
            int lastColumn = firstColumn + ColumnCount() - 1;
            if (lastColumn > XLHelper.MaxColumnNumber)
                lastColumn = XLHelper.MaxColumnNumber;

            int firstRow = RangeAddress.FirstAddress.RowNumber;
            int lastRow = firstRow + RowCount() - 1;
            if (lastRow > XLHelper.MaxRowNumber)
                lastRow = XLHelper.MaxRowNumber;

            var newRange = Worksheet.Range(firstRow, firstColumn, lastRow, lastColumn);
            return newRange.InsertColumnsBeforeInternal(onlyUsedCells, numberOfColumns, formatFromLeft, nullReturn);
        }

        public IXLRangeColumns InsertColumnsBefore(Int32 numberOfColumns)
        {
            return InsertColumnsBefore(numberOfColumns, false);
        }

        public IXLRangeColumns InsertColumnsBefore(Int32 numberOfColumns, Boolean expandRange)
        {
            var retVal = InsertColumnsBefore(false, numberOfColumns);
            // Adjust the range
            if (expandRange)
            {
                RangeAddress = new XLRangeAddress(
                    new XLAddress(Worksheet,
                                  RangeAddress.FirstAddress.RowNumber,
                                  RangeAddress.FirstAddress.ColumnNumber - numberOfColumns,
                                  RangeAddress.FirstAddress.FixedRow,
                                  RangeAddress.FirstAddress.FixedColumn),
                    new XLAddress(Worksheet,
                                  RangeAddress.LastAddress.RowNumber,
                                  RangeAddress.LastAddress.ColumnNumber,
                                  RangeAddress.LastAddress.FixedRow,
                                  RangeAddress.LastAddress.FixedColumn));
            }
            return retVal;
        }

        public IXLRangeColumns InsertColumnsBefore(Boolean onlyUsedCells, Int32 numberOfColumns, Boolean formatFromLeft = true)
        {
            return InsertColumnsBeforeInternal(onlyUsedCells, numberOfColumns, formatFromLeft);
        }

        public void InsertColumnsBeforeVoid(Boolean onlyUsedCells, Int32 numberOfColumns, Boolean formatFromLeft = true)
        {
            InsertColumnsBeforeInternal(onlyUsedCells, numberOfColumns, formatFromLeft, nullReturn: true);
        }

        private IXLRangeColumns InsertColumnsBeforeInternal(Boolean onlyUsedCells, Int32 numberOfColumns, Boolean formatFromLeft = true, Boolean nullReturn = false)
        {
            foreach (XLWorksheet ws in Worksheet.Workbook.WorksheetsInternal)
            {
                foreach (XLCell cell in ws.Internals.CellsCollection.GetCells(c => !XLHelper.IsNullOrWhiteSpace(c.FormulaA1)))
                    using (var asRange = AsRange())
                        cell.ShiftFormulaColumns(asRange, numberOfColumns);
            }


            var cellsDataValidations = new Dictionary<XLAddress, DataValidationToCopy>();
            var cellsToInsert = new Dictionary<IXLAddress, XLCell>();
            var cellsToDelete = new List<IXLAddress>();
            int firstColumn = RangeAddress.FirstAddress.ColumnNumber;
            int firstRow = RangeAddress.FirstAddress.RowNumber;
            int lastRow = RangeAddress.FirstAddress.RowNumber + RowCount() - 1;

            if (!onlyUsedCells)
            {
                int lastColumn = Worksheet.Internals.CellsCollection.MaxColumnUsed;
                if (lastColumn > 0)
                {
                    for (int co = lastColumn; co >= firstColumn; co--)
                    {
                        for (int ro = lastRow; ro >= firstRow; ro--)
                        {
                            var oldKey = new XLAddress(Worksheet, ro, co, false, false);
                            int newColumn = co + numberOfColumns;
                            var newKey = new XLAddress(Worksheet, ro, newColumn, false, false);
                            var oldCell = Worksheet.Internals.CellsCollection.GetCell(ro, co) ??
                                          Worksheet.Cell(oldKey);

                            var newCell = new XLCell(Worksheet, newKey, oldCell.GetStyleId());
                            newCell.CopyValuesFrom(oldCell);
                            newCell.FormulaA1 = oldCell.FormulaA1;
                            cellsToInsert.Add(newKey, newCell);
                            cellsToDelete.Add(oldKey);
                        }
                    }
                }
            }
            else
            {
                foreach (
                    XLCell c in
                        Worksheet.Internals.CellsCollection.GetCells(firstRow, firstColumn, lastRow,
                                                                     Worksheet.Internals.CellsCollection.MaxColumnUsed))
                {
                    int newColumn = c.Address.ColumnNumber + numberOfColumns;
                    var newKey = new XLAddress(Worksheet, c.Address.RowNumber, newColumn, false, false);
                    var newCell = new XLCell(Worksheet, newKey, c.GetStyleId());
                    newCell.CopyValuesFrom(c);
                    if (c.HasDataValidation)
                    {
                        cellsDataValidations.Add(newCell.Address,
                                                 new DataValidationToCopy
                                                 { DataValidation = c.DataValidation, SourceAddress = c.Address });
                        c.DataValidation.Clear();
                    }
                    newCell.FormulaA1 = c.FormulaA1;
                    cellsToInsert.Add(newKey, newCell);
                    cellsToDelete.Add(c.Address);
                }
            }

            cellsDataValidations.ForEach(kp =>
            {
                XLCell targetCell;
                if (!cellsToInsert.TryGetValue(kp.Key, out targetCell))
                    targetCell = Worksheet.Cell(kp.Key);

                targetCell.CopyDataValidation(Worksheet.Cell(kp.Value.SourceAddress), kp.Value.DataValidation);
            });

            cellsToDelete.ForEach(c => Worksheet.Internals.CellsCollection.Remove(c.RowNumber, c.ColumnNumber));
            cellsToInsert.ForEach(
                c => Worksheet.Internals.CellsCollection.Add(c.Key.RowNumber, c.Key.ColumnNumber, c.Value));
            //cellsDataValidations.ForEach(kp => Worksheet.Cell(kp.Key).CopyDataValidation(Worksheet.Cell(kp.Value.SourceAddress), kp.Value.DataValidation));

            Int32 firstRowReturn = RangeAddress.FirstAddress.RowNumber;
            Int32 lastRowReturn = RangeAddress.LastAddress.RowNumber;
            Int32 firstColumnReturn = RangeAddress.FirstAddress.ColumnNumber;
            Int32 lastColumnReturn = RangeAddress.FirstAddress.ColumnNumber + numberOfColumns - 1;

            Worksheet.BreakConditionalFormatsIntoCells(cellsToDelete.Except(cellsToInsert.Keys).ToList());
            using (var asRange = AsRange())
                Worksheet.NotifyRangeShiftedColumns(asRange, numberOfColumns);

            var rangeToReturn = Worksheet.Range(firstRowReturn, firstColumnReturn, lastRowReturn, lastColumnReturn);

            if (formatFromLeft && rangeToReturn.RangeAddress.FirstAddress.ColumnNumber > 1)
            {
                using (var firstColumnUsed = rangeToReturn.FirstColumn())
                {
                    using (var model = firstColumnUsed.ColumnLeft())
                    {
                        var modelFirstRow = model.FirstCellUsed(true);
                        var modelLastRow = model.LastCellUsed(true);
                        if (modelLastRow != null)
                        {
                            Int32 firstRoReturned = modelFirstRow.Address.RowNumber
                                                    - model.RangeAddress.FirstAddress.RowNumber + 1;
                            Int32 lastRoReturned = modelLastRow.Address.RowNumber
                                                   - model.RangeAddress.FirstAddress.RowNumber + 1;
                            for (Int32 ro = firstRoReturned; ro <= lastRoReturned; ro++)
                            {
                                rangeToReturn.Row(ro).Style = model.Cell(ro).Style;
                            }
                        }
                    }
                }
            }
            else
            {
                var lastRoUsed = rangeToReturn.LastRowUsed(true);
                if (lastRoUsed != null)
                {
                    Int32 lastRoReturned = lastRoUsed.RowNumber();
                    for (Int32 ro = 1; ro <= lastRoReturned; ro++)
                    {
                        var styleToUse = Worksheet.Internals.RowsCollection.ContainsKey(ro)
                                             ? Worksheet.Internals.RowsCollection[ro].Style
                                             : Worksheet.Style;
                        rangeToReturn.Row(ro).Style = styleToUse;
                    }

                }
            }

            if (nullReturn)
                return null;

            return rangeToReturn.Columns();
        }

        public IXLRangeRows InsertRowsBelow(Int32 numberOfRows)
        {
            return InsertRowsBelow(numberOfRows, true);
        }

        public IXLRangeRows InsertRowsBelow(Int32 numberOfRows, Boolean expandRange)
        {
            var retVal = InsertRowsBelow(false, numberOfRows);
            // Adjust the range
            if (expandRange)
            {
                RangeAddress = new XLRangeAddress(
                    new XLAddress(Worksheet,
                                  RangeAddress.FirstAddress.RowNumber,
                                  RangeAddress.FirstAddress.ColumnNumber,
                                  RangeAddress.FirstAddress.FixedRow,
                                  RangeAddress.FirstAddress.FixedColumn),
                    new XLAddress(Worksheet,
                                  RangeAddress.LastAddress.RowNumber + numberOfRows,
                                  RangeAddress.LastAddress.ColumnNumber,
                                  RangeAddress.LastAddress.FixedRow,
                                  RangeAddress.LastAddress.FixedColumn));
            }
            return retVal;
        }

        public IXLRangeRows InsertRowsBelow(Boolean onlyUsedCells, Int32 numberOfRows, Boolean formatFromAbove = true)
        {
            return InsertRowsBelowInternal(onlyUsedCells, numberOfRows, formatFromAbove, nullReturn: false);
        }

        public void InsertRowsBelowVoid(Boolean onlyUsedCells, Int32 numberOfRows, Boolean formatFromAbove = true)
        {
            InsertRowsBelowInternal(onlyUsedCells, numberOfRows, formatFromAbove, nullReturn: true);
        }

        private IXLRangeRows InsertRowsBelowInternal(Boolean onlyUsedCells, Int32 numberOfRows, Boolean formatFromAbove, Boolean nullReturn)
        {
            int rowCount = RowCount();
            int firstRow = RangeAddress.FirstAddress.RowNumber + rowCount;
            if (firstRow > XLHelper.MaxRowNumber)
                firstRow = XLHelper.MaxRowNumber;
            int lastRow = firstRow + RowCount() - 1;
            if (lastRow > XLHelper.MaxRowNumber)
                lastRow = XLHelper.MaxRowNumber;

            int firstColumn = RangeAddress.FirstAddress.ColumnNumber;
            int lastColumn = firstColumn + ColumnCount() - 1;
            if (lastColumn > XLHelper.MaxColumnNumber)
                lastColumn = XLHelper.MaxColumnNumber;

            var newRange = Worksheet.Range(firstRow, firstColumn, lastRow, lastColumn);
            return newRange.InsertRowsAboveInternal(onlyUsedCells, numberOfRows, formatFromAbove, nullReturn);
        }

        public IXLRangeRows InsertRowsAbove(Int32 numberOfRows)
        {
            return InsertRowsAbove(numberOfRows, false);
        }

        public IXLRangeRows InsertRowsAbove(Int32 numberOfRows, Boolean expandRange)
        {
            var retVal = InsertRowsAbove(false, numberOfRows);
            // Adjust the range
            if (expandRange)
            {
                RangeAddress = new XLRangeAddress(
                    new XLAddress(Worksheet,
                                  RangeAddress.FirstAddress.RowNumber - numberOfRows,
                                  RangeAddress.FirstAddress.ColumnNumber,
                                  RangeAddress.FirstAddress.FixedRow,
                                  RangeAddress.FirstAddress.FixedColumn),
                    new XLAddress(Worksheet,
                                  RangeAddress.LastAddress.RowNumber,
                                  RangeAddress.LastAddress.ColumnNumber,
                                  RangeAddress.LastAddress.FixedRow,
                                  RangeAddress.LastAddress.FixedColumn));
            }
            return retVal;
        }

        struct DataValidationToCopy
        {
            public XLAddress SourceAddress;
            public XLDataValidation DataValidation;
        }
        public void InsertRowsAboveVoid(Boolean onlyUsedCells, Int32 numberOfRows, Boolean formatFromAbove = true)
        {
            InsertRowsAboveInternal(onlyUsedCells, numberOfRows, formatFromAbove, nullReturn: true);
        }
        public IXLRangeRows InsertRowsAbove(Boolean onlyUsedCells, Int32 numberOfRows, Boolean formatFromAbove = true)
        {
            return InsertRowsAboveInternal(onlyUsedCells, numberOfRows, formatFromAbove, nullReturn: false);
        }

        private IXLRangeRows InsertRowsAboveInternal(Boolean onlyUsedCells, Int32 numberOfRows, Boolean formatFromAbove, Boolean nullReturn)
        {
            using (var asRange = AsRange())
                foreach (XLWorksheet ws in Worksheet.Workbook.WorksheetsInternal)
                {
                    foreach (XLCell cell in ws.Internals.CellsCollection.GetCells(c => !XLHelper.IsNullOrWhiteSpace(c.FormulaA1)))
                        cell.ShiftFormulaRows(asRange, numberOfRows);
                }

            var cellsToInsert = new Dictionary<IXLAddress, XLCell>();
            var cellsToDelete = new List<IXLAddress>();
            var cellsDataValidations = new Dictionary<XLAddress, DataValidationToCopy>();
            int firstRow = RangeAddress.FirstAddress.RowNumber;
            int firstColumn = RangeAddress.FirstAddress.ColumnNumber;
            int lastColumn = Math.Min(
                RangeAddress.FirstAddress.ColumnNumber + ColumnCount() - 1,
                Worksheet.Internals.CellsCollection.MaxColumnUsed);

            if (!onlyUsedCells)
            {
                int lastRow = Worksheet.Internals.CellsCollection.MaxRowUsed;
                if (lastRow > 0)
                {
                    for (int ro = lastRow; ro >= firstRow; ro--)
                    {
                        for (int co = lastColumn; co >= firstColumn; co--)
                        {
                            var oldKey = new XLAddress(Worksheet, ro, co, false, false);
                            int newRow = ro + numberOfRows;
                            var newKey = new XLAddress(Worksheet, newRow, co, false, false);
                            var oldCell = Worksheet.Internals.CellsCollection.GetCell(ro, co);
                            if (oldCell != null)
                            {
                                var newCell = new XLCell(Worksheet, newKey, oldCell.GetStyleId());
                                newCell.CopyValuesFrom(oldCell);
                                newCell.FormulaA1 = oldCell.FormulaA1;
                                cellsToInsert.Add(newKey, newCell);
                                cellsToDelete.Add(oldKey);
                            }
                        }
                    }
                }
            }
            else
            {
                foreach (
                    XLCell c in
                        Worksheet.Internals.CellsCollection.GetCells(firstRow, firstColumn,
                                                                     Worksheet.Internals.CellsCollection.MaxRowUsed,
                                                                     lastColumn))
                {
                    int newRow = c.Address.RowNumber + numberOfRows;
                    var newKey = new XLAddress(Worksheet, newRow, c.Address.ColumnNumber, false, false);
                    var newCell = new XLCell(Worksheet, newKey, c.GetStyleId());
                    newCell.CopyValuesFrom(c);
                    if (c.HasDataValidation)
                    {
                        cellsDataValidations.Add(newCell.Address,
                                                 new DataValidationToCopy
                                                 { DataValidation = c.DataValidation, SourceAddress = c.Address });
                        c.DataValidation.Clear();
                    }
                    newCell.FormulaA1 = c.FormulaA1;
                    cellsToInsert.Add(newKey, newCell);
                    cellsToDelete.Add(c.Address);

                }
            }

            cellsDataValidations
                .ForEach(kp =>
                {
                    XLCell targetCell;
                    if (!cellsToInsert.TryGetValue(kp.Key, out targetCell))
                        targetCell = Worksheet.Cell(kp.Key);

                    targetCell.CopyDataValidation(
                        Worksheet.Cell(kp.Value.SourceAddress), kp.Value.DataValidation);
                });

            cellsToDelete.ForEach(c => Worksheet.Internals.CellsCollection.Remove(c.RowNumber, c.ColumnNumber));
            cellsToInsert.ForEach(c => Worksheet.Internals.CellsCollection.Add(c.Key.RowNumber, c.Key.ColumnNumber, c.Value));


            Int32 firstRowReturn = RangeAddress.FirstAddress.RowNumber;
            Int32 lastRowReturn = RangeAddress.FirstAddress.RowNumber + numberOfRows - 1;
            Int32 firstColumnReturn = RangeAddress.FirstAddress.ColumnNumber;
            Int32 lastColumnReturn = RangeAddress.LastAddress.ColumnNumber;

            Worksheet.BreakConditionalFormatsIntoCells(cellsToDelete.Except(cellsToInsert.Keys).ToList());
            using (var asRange = AsRange())
                Worksheet.NotifyRangeShiftedRows(asRange, numberOfRows);

            var rangeToReturn = Worksheet.Range(firstRowReturn, firstColumnReturn, lastRowReturn, lastColumnReturn);

            if (formatFromAbove && rangeToReturn.RangeAddress.FirstAddress.RowNumber > 1)
            {
                using (var fr = rangeToReturn.FirstRow())
                {
                    using (var model = fr.RowAbove())
                    {
                        var modelFirstColumn = model.FirstCellUsed(true);
                        var modelLastColumn = model.LastCellUsed(true);
                        if (modelFirstColumn != null && modelLastColumn != null)
                        {
                            Int32 firstCoReturned = modelFirstColumn.Address.ColumnNumber
                                                    - model.RangeAddress.FirstAddress.ColumnNumber + 1;
                            Int32 lastCoReturned = modelLastColumn.Address.ColumnNumber
                                                   - model.RangeAddress.FirstAddress.ColumnNumber + 1;
                            for (Int32 co = firstCoReturned; co <= lastCoReturned; co++)
                            {
                                rangeToReturn.Column(co).Style = model.Cell(co).Style;
                            }
                        }
                    }
                }
            }
            else
            {
                var lastCoUsed = rangeToReturn.LastColumnUsed(true);
                if (lastCoUsed != null)
                {
                    Int32 lastCoReturned = lastCoUsed.ColumnNumber();
                    for (Int32 co = 1; co <= lastCoReturned; co++)
                    {
                        var styleToUse = Worksheet.Internals.ColumnsCollection.ContainsKey(co)
                                             ? Worksheet.Internals.ColumnsCollection[co].Style
                                             : Worksheet.Style;
                        rangeToReturn.Column(co).Style = styleToUse;
                    }
                }
            }

            // Skip calling .Rows() for performance reasons if required.
            if (nullReturn)
                return null;

            return rangeToReturn.Rows();
        }

        private void ClearMerged()
        {
            var mergeToDelete = Worksheet.Internals.MergedRanges.Where(Intersects).ToList();
            mergeToDelete.ForEach(m => Worksheet.Internals.MergedRanges.Remove(m));
        }

        public Boolean Contains(IXLCell cell)
        {
            return Contains(cell.Address as XLAddress);
        }

        public bool Contains(XLAddress first, XLAddress last)
        {
            return Contains(first) && Contains(last);
        }

        public bool Contains(XLAddress address)
        {
            return RangeAddress.FirstAddress.RowNumber <= address.RowNumber &&
                   address.RowNumber <= RangeAddress.LastAddress.RowNumber &&
                   RangeAddress.FirstAddress.ColumnNumber <= address.ColumnNumber &&
                   address.ColumnNumber <= RangeAddress.LastAddress.ColumnNumber;
        }

        public void Delete(XLShiftDeletedCells shiftDeleteCells)
        {
            int numberOfRows = RowCount();
            int numberOfColumns = ColumnCount();
            IXLRange shiftedRangeFormula = Worksheet.Range(
                RangeAddress.FirstAddress.RowNumber,
                RangeAddress.FirstAddress.ColumnNumber,
                RangeAddress.LastAddress.RowNumber,
                RangeAddress.LastAddress.ColumnNumber);


            foreach (
                XLCell cell in
                    Worksheet.Workbook.Worksheets.Cast<XLWorksheet>().SelectMany(
                        xlWorksheet => (xlWorksheet).Internals.CellsCollection.GetCells(
                            c => !XLHelper.IsNullOrWhiteSpace(c.FormulaA1))))
            {
                if (shiftDeleteCells == XLShiftDeletedCells.ShiftCellsUp)
                    cell.ShiftFormulaRows((XLRange)shiftedRangeFormula, numberOfRows * -1);
                else
                    cell.ShiftFormulaColumns((XLRange)shiftedRangeFormula, numberOfColumns * -1);
            }

            // Range to shift...
            var cellsToInsert = new Dictionary<IXLAddress, XLCell>();
            //var cellsDataValidations = new Dictionary<XLAddress, DataValidationToCopy>();
            var cellsToDelete = new List<IXLAddress>();
            var shiftLeftQuery = Worksheet.Internals.CellsCollection.GetCells(
                RangeAddress.FirstAddress.RowNumber,
                RangeAddress.FirstAddress.ColumnNumber,
                RangeAddress.LastAddress.RowNumber,
                Worksheet.Internals.CellsCollection.MaxColumnUsed);

            var shiftUpQuery = Worksheet.Internals.CellsCollection.GetCells(
                RangeAddress.FirstAddress.RowNumber,
                RangeAddress.FirstAddress.ColumnNumber,
                Worksheet.Internals.CellsCollection.MaxRowUsed,
                RangeAddress.LastAddress.ColumnNumber);


            int columnModifier = shiftDeleteCells == XLShiftDeletedCells.ShiftCellsLeft ? ColumnCount() : 0;
            int rowModifier = shiftDeleteCells == XLShiftDeletedCells.ShiftCellsUp ? RowCount() : 0;
            var cellsQuery = shiftDeleteCells == XLShiftDeletedCells.ShiftCellsLeft ? shiftLeftQuery : shiftUpQuery;
            foreach (XLCell c in cellsQuery)
            {
                var newKey = new XLAddress(Worksheet, c.Address.RowNumber - rowModifier,
                                           c.Address.ColumnNumber - columnModifier,
                                           false, false);
                var newCell = new XLCell(Worksheet, newKey, c.GetStyleId());
                newCell.CopyValuesFrom(c);
                newCell.FormulaA1 = c.FormulaA1;
                cellsToDelete.Add(c.Address);

                bool canInsert = shiftDeleteCells == XLShiftDeletedCells.ShiftCellsLeft
                                     ? c.Address.ColumnNumber > RangeAddress.LastAddress.ColumnNumber
                                     : c.Address.RowNumber > RangeAddress.LastAddress.RowNumber;

                if (canInsert)
                    cellsToInsert.Add(newKey, newCell);
            }


            cellsToDelete.ForEach(c => Worksheet.Internals.CellsCollection.Remove(c.RowNumber, c.ColumnNumber));
            cellsToInsert.ForEach(
                c => Worksheet.Internals.CellsCollection.Add(c.Key.RowNumber, c.Key.ColumnNumber, c.Value));

            var mergesToRemove = Worksheet.Internals.MergedRanges.Where(Contains).ToList();
            mergesToRemove.ForEach(r => Worksheet.Internals.MergedRanges.Remove(r));

            var hyperlinksToRemove = Worksheet.Hyperlinks.Where(hl => Contains(hl.Cell.AsRange())).ToList();
            hyperlinksToRemove.ForEach(hl => Worksheet.Hyperlinks.Delete(hl));

            Worksheet.BreakConditionalFormatsIntoCells(cellsToDelete.Except(cellsToInsert.Keys).ToList());
            using (var shiftedRange = AsRange())
            {
                if (shiftDeleteCells == XLShiftDeletedCells.ShiftCellsUp)
                    Worksheet.NotifyRangeShiftedRows(shiftedRange, rowModifier * -1);
                else
                    Worksheet.NotifyRangeShiftedColumns(shiftedRange, columnModifier * -1);
            }
        }

        public override string ToString()
        {
            return String.Format("{0}!{1}:{2}", Worksheet.Name.WrapSheetNameInQuotesIfRequired(), RangeAddress.FirstAddress, RangeAddress.LastAddress);
        }

        protected void ShiftColumns(IXLRangeAddress thisRangeAddress, XLRange shiftedRange, int columnsShifted)
        {
            if (thisRangeAddress.IsInvalid || shiftedRange.RangeAddress.IsInvalid) return;

            if ((columnsShifted < 0
                 // all columns
                 &&
                 thisRangeAddress.FirstAddress.ColumnNumber >= shiftedRange.RangeAddress.FirstAddress.ColumnNumber
                 &&
                 thisRangeAddress.LastAddress.ColumnNumber <=
                 shiftedRange.RangeAddress.FirstAddress.ColumnNumber - columnsShifted
                 // all rows
                 && thisRangeAddress.FirstAddress.RowNumber >= shiftedRange.RangeAddress.FirstAddress.RowNumber
                 && thisRangeAddress.LastAddress.RowNumber <= shiftedRange.RangeAddress.LastAddress.RowNumber
                ) || (
                         shiftedRange.RangeAddress.FirstAddress.ColumnNumber <=
                         thisRangeAddress.FirstAddress.ColumnNumber
                         &&
                         shiftedRange.RangeAddress.FirstAddress.RowNumber <= thisRangeAddress.FirstAddress.RowNumber
                         &&
                         shiftedRange.RangeAddress.LastAddress.RowNumber >= thisRangeAddress.LastAddress.RowNumber
                         && shiftedRange.ColumnCount() >
                         (thisRangeAddress.LastAddress.ColumnNumber - thisRangeAddress.FirstAddress.ColumnNumber + 1)
                         +
                         (thisRangeAddress.FirstAddress.ColumnNumber -
                          shiftedRange.RangeAddress.FirstAddress.ColumnNumber)))
                thisRangeAddress.IsInvalid = true;
            else
            {
                if (shiftedRange.RangeAddress.FirstAddress.RowNumber <= thisRangeAddress.FirstAddress.RowNumber
                    && shiftedRange.RangeAddress.LastAddress.RowNumber >= thisRangeAddress.LastAddress.RowNumber)
                {
                    if (
                        (shiftedRange.RangeAddress.FirstAddress.ColumnNumber <=
                         thisRangeAddress.FirstAddress.ColumnNumber &&
                         columnsShifted > 0)
                        ||
                        (shiftedRange.RangeAddress.FirstAddress.ColumnNumber <
                         thisRangeAddress.FirstAddress.ColumnNumber &&
                         columnsShifted < 0)
                        )
                    {
                        thisRangeAddress.FirstAddress = new XLAddress(Worksheet,
                                                                      thisRangeAddress.FirstAddress.RowNumber,
                                                                      thisRangeAddress.FirstAddress.ColumnNumber +
                                                                      columnsShifted,
                                                                      thisRangeAddress.FirstAddress.FixedRow,
                                                                      thisRangeAddress.FirstAddress.FixedColumn);
                    }

                    if (shiftedRange.RangeAddress.FirstAddress.ColumnNumber <=
                        thisRangeAddress.LastAddress.ColumnNumber)
                    {
                        thisRangeAddress.LastAddress = new XLAddress(Worksheet,
                                                                     thisRangeAddress.LastAddress.RowNumber,
                                                                     thisRangeAddress.LastAddress.ColumnNumber +
                                                                     columnsShifted,
                                                                     thisRangeAddress.LastAddress.FixedRow,
                                                                     thisRangeAddress.LastAddress.FixedColumn);
                    }
                }
            }
        }

        protected void ShiftRows(IXLRangeAddress thisRangeAddress, XLRange shiftedRange, int rowsShifted)
        {
            if (thisRangeAddress.IsInvalid || shiftedRange.RangeAddress.IsInvalid) return;

            if ((rowsShifted < 0
                 // all columns
                 && thisRangeAddress.FirstAddress.ColumnNumber >= shiftedRange.RangeAddress.FirstAddress.ColumnNumber
                 && thisRangeAddress.LastAddress.ColumnNumber <= shiftedRange.RangeAddress.FirstAddress.ColumnNumber
                 // all rows
                 && thisRangeAddress.FirstAddress.RowNumber >= shiftedRange.RangeAddress.FirstAddress.RowNumber
                 && thisRangeAddress.LastAddress.RowNumber <= shiftedRange.RangeAddress.LastAddress.RowNumber - rowsShifted
                ) || (
                         shiftedRange.RangeAddress.FirstAddress.RowNumber <= thisRangeAddress.FirstAddress.RowNumber
                         && shiftedRange.RangeAddress.FirstAddress.ColumnNumber <= thisRangeAddress.FirstAddress.ColumnNumber
                         && shiftedRange.RangeAddress.LastAddress.ColumnNumber >= thisRangeAddress.LastAddress.ColumnNumber
                         && shiftedRange.RowCount() >
                            (thisRangeAddress.LastAddress.RowNumber - thisRangeAddress.FirstAddress.RowNumber + 1)
                            + (thisRangeAddress.FirstAddress.RowNumber - shiftedRange.RangeAddress.FirstAddress.RowNumber)))
                thisRangeAddress.IsInvalid = true;
            else
            {
                if (shiftedRange.RangeAddress.FirstAddress.ColumnNumber <= thisRangeAddress.FirstAddress.ColumnNumber
                    && shiftedRange.RangeAddress.LastAddress.ColumnNumber >= thisRangeAddress.LastAddress.ColumnNumber)
                {
                    if (
                        (shiftedRange.RangeAddress.FirstAddress.RowNumber <= thisRangeAddress.FirstAddress.RowNumber && rowsShifted > 0)
                        || (shiftedRange.RangeAddress.FirstAddress.RowNumber < thisRangeAddress.FirstAddress.RowNumber && rowsShifted < 0)
                        )
                    {
                        thisRangeAddress.FirstAddress = new XLAddress(Worksheet,
                                                                      thisRangeAddress.FirstAddress.RowNumber +
                                                                      rowsShifted,
                                                                      thisRangeAddress.FirstAddress.ColumnNumber,
                                                                      thisRangeAddress.FirstAddress.FixedRow,
                                                                      thisRangeAddress.FirstAddress.FixedColumn);
                    }

                    if (shiftedRange.RangeAddress.FirstAddress.RowNumber <= thisRangeAddress.LastAddress.RowNumber)
                    {
                        thisRangeAddress.LastAddress = new XLAddress(Worksheet,
                                                                     thisRangeAddress.LastAddress.RowNumber +
                                                                     rowsShifted,
                                                                     thisRangeAddress.LastAddress.ColumnNumber,
                                                                     thisRangeAddress.LastAddress.FixedRow,
                                                                     thisRangeAddress.LastAddress.FixedColumn);
                    }
                }
            }
        }

        public IXLRange RangeUsed()
        {
            return RangeUsed(false);
        }

        public IXLRange RangeUsed(bool includeFormats)
        {
            var firstCell = FirstCellUsed(includeFormats);
            if (firstCell == null)
                return null;
            var lastCell = LastCellUsed(includeFormats);
            return Worksheet.Range(firstCell, lastCell);
        }

        public virtual void CopyTo(IXLRangeBase target)
        {
            CopyTo(target.FirstCell());
        }

        public virtual void CopyTo(IXLCell target)
        {
            target.Value = this;
        }

        //public IXLChart CreateChart(Int32 firstRow, Int32 firstColumn, Int32 lastRow, Int32 lastColumn)
        //{
        //    IXLChart chart = new XLChartWorksheet;
        //    chart.FirstRow = firstRow;
        //    chart.LastRow = lastRow;
        //    chart.LastColumn = lastColumn;
        //    chart.FirstColumn = firstColumn;
        //    Worksheet.Charts.Add(chart);
        //    return chart;
        //}


        IXLPivotTable IXLRangeBase.CreatePivotTable(IXLCell targetCell)
        {
            return CreatePivotTable(targetCell);
        }
        IXLPivotTable IXLRangeBase.CreatePivotTable(IXLCell targetCell, String name)
        {
            return CreatePivotTable(targetCell, name);
        }


        public XLPivotTable CreatePivotTable(IXLCell targetCell)
        {
            return CreatePivotTable(targetCell, Guid.NewGuid().ToString());
        }

        public XLPivotTable CreatePivotTable(IXLCell targetCell, String name)
        {
            return (XLPivotTable)Worksheet.PivotTables.AddNew(name, targetCell, AsRange());
        }

        public IXLAutoFilter SetAutoFilter()
        {
            using (var asRange = AsRange())
                return Worksheet.AutoFilter.Set(asRange);
        }

        #region Sort

        public IXLSortElements SortRows
        {
            get { return _sortRows ?? (_sortRows = new XLSortElements()); }
        }

        public IXLSortElements SortColumns
        {
            get { return _sortColumns ?? (_sortColumns = new XLSortElements()); }
        }

        public IXLRangeBase Sort()
        {
            if (!SortColumns.Any())
            {
                String columnsToSortBy = String.Empty;
                Int32 maxColumn = ColumnCount();
                if (maxColumn == XLHelper.MaxColumnNumber)
                    maxColumn = LastCellUsed(true).Address.ColumnNumber;
                for (int i = 1; i <= maxColumn; i++)
                {
                    columnsToSortBy += i + ",";
                }
                columnsToSortBy = columnsToSortBy.Substring(0, columnsToSortBy.Length - 1);
                return Sort(columnsToSortBy);
            }

            SortRangeRows();
            return this;
        }

        public IXLRangeBase Sort(String columnsToSortBy, XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true)
        {
            SortColumns.Clear();
            if (XLHelper.IsNullOrWhiteSpace(columnsToSortBy))
            {
                columnsToSortBy = String.Empty;
                Int32 maxColumn = ColumnCount();
                if (maxColumn == XLHelper.MaxColumnNumber)
                    maxColumn = LastCellUsed(true).Address.ColumnNumber;
                for (int i = 1; i <= maxColumn; i++)
                {
                    columnsToSortBy += i + ",";
                }
                columnsToSortBy = columnsToSortBy.Substring(0, columnsToSortBy.Length - 1);
            }

            foreach (string coPairTrimmed in columnsToSortBy.Split(',').Select(coPair => coPair.Trim()))
            {
                String coString;
                String order;
                if (coPairTrimmed.Contains(' '))
                {
                    var pair = coPairTrimmed.Split(' ');
                    coString = pair[0];
                    order = pair[1];
                }
                else
                {
                    coString = coPairTrimmed;
                    order = sortOrder == XLSortOrder.Ascending ? "ASC" : "DESC";
                }

                Int32 co;
                if (!Int32.TryParse(coString, out co))
                    co = XLHelper.GetColumnNumberFromLetter(coString);

                SortColumns.Add(co, String.Compare(order, "ASC", true) == 0 ? XLSortOrder.Ascending : XLSortOrder.Descending, ignoreBlanks, matchCase);
            }

            SortRangeRows();
            return this;
        }

        public IXLRangeBase Sort(Int32 columnToSortBy, XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true)
        {
            return Sort(columnToSortBy.ToString(), sortOrder, matchCase, ignoreBlanks);
        }

        public IXLRangeBase SortLeftToRight(XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true)
        {
            SortRows.Clear();
            Int32 maxColumn = ColumnCount();
            if (maxColumn == XLHelper.MaxColumnNumber)
                maxColumn = LastCellUsed(true).Address.ColumnNumber;

            for (int i = 1; i <= maxColumn; i++)
            {
                SortRows.Add(i, sortOrder, ignoreBlanks, matchCase);
            }

            SortRangeColumns();
            return this;
        }


        #region Sort Rows

        private void SortRangeRows()
        {
            Int32 maxRow = RowCount();
            if (maxRow == XLHelper.MaxRowNumber)
                maxRow = LastCellUsed(true).Address.RowNumber;

            SortingRangeRows(1, maxRow);
        }

        private void SwapRows(Int32 row1, Int32 row2)
        {
            int row1InWs = RangeAddress.FirstAddress.RowNumber + row1 - 1;
            int row2InWs = RangeAddress.FirstAddress.RowNumber + row2 - 1;

            Int32 firstColumn = RangeAddress.FirstAddress.ColumnNumber;
            Int32 lastColumn = RangeAddress.LastAddress.ColumnNumber;

            var range1Sp1 = new XLSheetPoint(row1InWs, firstColumn);
            var range1Sp2 = new XLSheetPoint(row1InWs, lastColumn);
            var range2Sp1 = new XLSheetPoint(row2InWs, firstColumn);
            var range2Sp2 = new XLSheetPoint(row2InWs, lastColumn);

            Worksheet.Internals.CellsCollection.SwapRanges(new XLSheetRange(range1Sp1, range1Sp2),
                                                           new XLSheetRange(range2Sp1, range2Sp2), Worksheet);
        }

        private int SortRangeRows(int begPoint, int endPoint)
        {
            int pivot = begPoint;
            int m = begPoint + 1;
            int n = endPoint;
            while ((m < endPoint) && RowQuick(pivot).CompareTo(RowQuick(m), SortColumns) >= 0)
                m++;

            while (n > begPoint && RowQuick(pivot).CompareTo(RowQuick(n), SortColumns) <= 0)
                n--;
            while (m < n)
            {
                SwapRows(m, n);

                while (m < endPoint && RowQuick(pivot).CompareTo(RowQuick(m), SortColumns) >= 0)
                    m++;

                while (n > begPoint && RowQuick(pivot).CompareTo(RowQuick(n), SortColumns) <= 0)
                    n--;
            }
            if (pivot != n)
                SwapRows(n, pivot);
            return n;
        }

        private void SortingRangeRows(int beg, int end)
        {
            if (end == beg)
                return;
            int pivot = SortRangeRows(beg, end);
            if (pivot > beg)
                SortingRangeRows(beg, pivot - 1);
            if (pivot < end)
                SortingRangeRows(pivot + 1, end);
        }

        #endregion

        #region Sort Columns

        private void SortRangeColumns()
        {
            Int32 maxColumn = ColumnCount();
            if (maxColumn == XLHelper.MaxColumnNumber)
                maxColumn = LastCellUsed(true).Address.ColumnNumber;
            SortingRangeColumns(1, maxColumn);
        }

        private void SwapColumns(Int32 column1, Int32 column2)
        {
            int col1InWs = RangeAddress.FirstAddress.ColumnNumber + column1 - 1;
            int col2InWs = RangeAddress.FirstAddress.ColumnNumber + column2 - 1;

            Int32 firstRow = RangeAddress.FirstAddress.RowNumber;
            Int32 lastRow = RangeAddress.LastAddress.RowNumber;

            var range1Sp1 = new XLSheetPoint(firstRow, col1InWs);
            var range1Sp2 = new XLSheetPoint(lastRow, col1InWs);
            var range2Sp1 = new XLSheetPoint(firstRow, col2InWs);
            var range2Sp2 = new XLSheetPoint(lastRow, col2InWs);

            Worksheet.Internals.CellsCollection.SwapRanges(new XLSheetRange(range1Sp1, range1Sp2),
                                                           new XLSheetRange(range2Sp1, range2Sp2), Worksheet);
        }

        private int SortRangeColumns(int begPoint, int endPoint)
        {
            int pivot = begPoint;
            int m = begPoint + 1;
            int n = endPoint;
            while ((m < endPoint) && ColumnQuick(pivot).CompareTo((ColumnQuick(m)), SortRows) >= 0)
                m++;

            while ((n > begPoint) && ((ColumnQuick(pivot)).CompareTo((ColumnQuick(n)), SortRows) <= 0))
                n--;
            while (m < n)
            {
                SwapColumns(m, n);

                while ((m < endPoint) && (ColumnQuick(pivot)).CompareTo((ColumnQuick(m)), SortRows) >= 0)
                    m++;

                while ((n > begPoint) && (ColumnQuick(pivot)).CompareTo((ColumnQuick(n)), SortRows) <= 0)
                    n--;
            }
            if (pivot != n)
                SwapColumns(n, pivot);
            return n;
        }

        private void SortingRangeColumns(int beg, int end)
        {
            if (end == beg)
                return;
            int pivot = SortRangeColumns(beg, end);
            if (pivot > beg)
                SortingRangeColumns(beg, pivot - 1);
            if (pivot < end)
                SortingRangeColumns(pivot + 1, end);
        }

        #endregion

        #endregion

        public XLRangeColumn ColumnQuick(Int32 column)
        {
            var firstCellAddress = new XLAddress(Worksheet,
                                                 RangeAddress.FirstAddress.RowNumber,
                                                 RangeAddress.FirstAddress.ColumnNumber + column - 1,
                                                 false,
                                                 false);
            var lastCellAddress = new XLAddress(Worksheet,
                                                RangeAddress.LastAddress.RowNumber,
                                                RangeAddress.FirstAddress.ColumnNumber + column - 1,
                                                false,
                                                false);
            return new XLRangeColumn(
                new XLRangeParameters(new XLRangeAddress(firstCellAddress, lastCellAddress), Worksheet.Style), true);
        }

        public XLRangeRow RowQuick(Int32 row)
        {
            var firstCellAddress = new XLAddress(Worksheet,
                                                 RangeAddress.FirstAddress.RowNumber + row - 1,
                                                 RangeAddress.FirstAddress.ColumnNumber,
                                                 false,
                                                 false);
            var lastCellAddress = new XLAddress(Worksheet,
                                                RangeAddress.FirstAddress.RowNumber + row - 1,
                                                RangeAddress.LastAddress.ColumnNumber,
                                                false,
                                                false);
            return new XLRangeRow(
                new XLRangeParameters(new XLRangeAddress(firstCellAddress, lastCellAddress), Worksheet.Style), true);
        }

        public void Dispose()
        {
            if (_shiftedRowsAction != null)
            {
                RangeAddress.Worksheet.RangeShiftedRows.Remove(_shiftedRowsAction);
                _shiftedRowsAction = null;
            }

            if (_shiftedColumnsAction != null)
            {
                RangeAddress.Worksheet.RangeShiftedColumns.Remove(_shiftedColumnsAction);
                _shiftedColumnsAction = null;
            }
        }

        public IXLDataValidation SetDataValidation()
        {
            return DataValidation;
        }

        public IXLConditionalFormat AddConditionalFormat()
        {
            using (var asRange = AsRange())
            {
                var cf = new XLConditionalFormat(asRange);
                Worksheet.ConditionalFormats.Add(cf);
                return cf;
            }
        }


        internal IXLConditionalFormat AddConditionalFormat(IXLConditionalFormat source)
        {
            using (var asRange = AsRange())
            {
                var cf = new XLConditionalFormat(asRange);
                cf.CopyFrom(source);
                Worksheet.ConditionalFormats.Add(cf);
                return cf;
            }
        }

        public void Select()
        {
            Worksheet.SelectedRanges.Add(AsRange());
        }
    }
}
