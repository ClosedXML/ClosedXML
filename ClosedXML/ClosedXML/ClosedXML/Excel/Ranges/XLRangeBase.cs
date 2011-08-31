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

        protected XLRangeBase(XLRangeAddress rangeAddress)
        {
            RangeAddress = rangeAddress;
        }

        #endregion

        #region Public properties

        public XLRangeAddress RangeAddress { get; protected set; }

        public XLWorksheet Worksheet
        {
            get { return RangeAddress.Worksheet; }
        }

        public XLDataValidation DataValidation
        {
            get
            {
                var thisRange = AsRange();
                if (Worksheet.DataValidations.ContainsSingle(thisRange))
                {
                    return
                        Worksheet.DataValidations.Where(dv => dv.Ranges.Contains(thisRange)).Single() as
                        XLDataValidation;
                }
                var dvEmpty = new List<IXLDataValidation>();
                foreach (IXLDataValidation dv in Worksheet.DataValidations)
                {
                    foreach (IXLRange dvRange in dv.Ranges.Where(dvRange => dvRange.Intersects(this)))
                    {
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
                                    dv.Ranges.Add(Worksheet.Column(column.ColumnNumber()).Column(
                                        dvStart, 
                                        thisStart - 1));
                                    dv.Ranges.Add(Worksheet.Column(column.ColumnNumber()).Column(
                                        thisEnd + 1,
                                        dvEnd));
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
                                            dv.Ranges.Add(Worksheet.Column(column.ColumnNumber()).Column(coStart, coEnd));
                                    }
                                }
                            }
                            else
                            {
                                dv.Ranges.Add(column);
                            }
                        }

                        if (dv.Ranges.Count() == 0)
                            dvEmpty.Add(dv);
                    }
                }

                dvEmpty.ForEach(dv => Worksheet.DataValidations.Delete(dv));

                var newRanges = new XLRanges {AsRange()};
                var dataValidation = new XLDataValidation(newRanges);

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
            set { Cells().ForEach(c => c.FormulaA1 = value); }
        }

        public String FormulaR1C1
        {
            set { Cells().ForEach(c => c.FormulaR1C1 = value); }
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

        IXLDataValidation IXLRangeBase.DataValidation
        {
            get { return DataValidation; }
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
                var retVal = new XLRanges {AsRange()};
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
            var cells = new XLCells(false, false) {RangeAddress};
            return cells;
        }

        public IXLCells Cells(String cells)
        {
            return Ranges(cells).Cells();
        }

        public IXLCells CellsUsed()
        {
            var cells = new XLCells(true, false) {RangeAddress};
            return cells;
        }

        IXLCells IXLRangeBase.CellsUsed(Boolean includeFormats)
        {
            return CellsUsed(includeFormats);
        }

        public IXLRange Merge()
        {
            string tAddress = RangeAddress.ToString();
            Boolean foundOne =
                Worksheet.Internals.MergedRanges.Select(m => m.RangeAddress.ToString()).Any(
                    mAddress => mAddress == tAddress);

            var asRange = AsRange();
            if (!foundOne)
                Worksheet.Internals.MergedRanges.Add(asRange);

            // Call every cell in the merge to make sure they exist
            asRange.Cells().ForEach(c => { });

            return asRange;
        }

        public IXLRange Unmerge()
        {
            string tAddress = RangeAddress.ToString();
            if (
                Worksheet.Internals.MergedRanges.Select(m => m.RangeAddress.ToString()).Any(
                    mAddress => mAddress == tAddress))
                Worksheet.Internals.MergedRanges.Remove(AsRange());

            return AsRange();
        }

        public IXLRangeBase Clear(XLClearOptions clearOptions = XLClearOptions.ContentsAndFormats)
        {
            var includeFormats = clearOptions == XLClearOptions.Formats ||
                                 clearOptions == XLClearOptions.ContentsAndFormats;
            foreach (var cell in CellsUsed(includeFormats))
            {
                cell.Clear(clearOptions);
            }

            if (includeFormats)
            {
                ClearMerged();

                var hyperlinksToRemove = Worksheet.Hyperlinks.Where(hl => Contains(hl.Cell.AsRange())).ToList();
                hyperlinksToRemove.ForEach(hl => Worksheet.Hyperlinks.Delete(hl));
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

        public void DeleteComments() {
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
            return Intersects(Range(rangeAddress));
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

        public virtual IXLRange AsRange()
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
                var namedRange = namedRanges.Where(nr => String.Compare(nr.Name, rangeName, true) == 0).Single();
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
            return CellsUsed().Any(c => c.IsMerged());
        }

        public Boolean IsEmpty()
        {
            return !CellsUsed().Any() || CellsUsed().Any(c => c.IsEmpty());
        }

        public Boolean IsEmpty(Boolean includeFormats)
        {
            return !CellsUsed(includeFormats).Any<XLCell>() ||
                   CellsUsed(includeFormats).Any<XLCell>(c => c.IsEmpty(includeFormats));
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
            return FirstCellUsed(false);
        }

        public XLCell FirstCellUsed(Boolean includeFormats)
        {
            var cellsUsed = CellsUsed(includeFormats);

            if (!cellsUsed.Any<XLCell>())
                return null;
            int firstRow = cellsUsed.Min<XLCell>(c => c.Address.RowNumber);
            int firstColumn = cellsUsed.Min<XLCell>(c => c.Address.ColumnNumber);
            return Worksheet.Cell(firstRow, firstColumn);
        }

        public XLCell LastCellUsed()
        {
            return LastCellUsed(false);
        }

        public XLCell LastCellUsed(Boolean includeFormats)
        {
            var cellsUsed = CellsUsed(includeFormats);
            if (!cellsUsed.Any<XLCell>())
                return null;

            int lastRow = cellsUsed.Max<XLCell>(c => c.Address.RowNumber);
            int lastColumn = cellsUsed.Max<XLCell>(c => c.Address.ColumnNumber);
            return Worksheet.Cell(lastRow, lastColumn);
        }

        public XLCell Cell(Int32 row, Int32 column)
        {
            return Cell(new XLAddress(Worksheet, row, column, false, false));
        }

        public XLCell Cell(String cellAddressInRange)
        {

            if (ExcelHelper.IsValidA1Address(cellAddressInRange))
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
            var absoluteAddress = cellAddressInRange + RangeAddress.FirstAddress - 1;

            if (absoluteAddress.RowNumber <= 0 || absoluteAddress.RowNumber > ExcelHelper.MaxRowNumber)
            {
                throw new IndexOutOfRangeException(String.Format("Row number must be between 1 and {0}",
                                                                 ExcelHelper.MaxRowNumber));
            }

            if (absoluteAddress.ColumnNumber <= 0 || absoluteAddress.ColumnNumber > ExcelHelper.MaxColumnNumber)
            {
                throw new IndexOutOfRangeException(String.Format("Column number must be between 1 and {0}",
                                                                 ExcelHelper.MaxColumnNumber));
            }

            var cell = Worksheet.Internals.CellsCollection.GetCell(absoluteAddress.RowNumber,
                                                                   absoluteAddress.ColumnNumber);

            if (cell != null)
                return cell;

            //var style = Style;
            Int32 styleId = GetStyleId();
            Int32 worksheetStyleId = Worksheet.GetStyleId();
            
            if (styleId == worksheetStyleId)
            {
                XLRow row;
                XLColumn column;
                if (Worksheet.Internals.RowsCollection.TryGetValue(absoluteAddress.RowNumber, out row)
                    && row.GetStyleId() == worksheetStyleId)
                    styleId = row.GetStyleId();
                else if (Worksheet.Internals.ColumnsCollection.TryGetValue(absoluteAddress.ColumnNumber, out column)
                    && column.GetStyleId() == worksheetStyleId)
                    styleId = column.GetStyleId();
                //if (Worksheet.Internals.RowsCollection.ContainsKey(absoluteAddress.RowNumber)
                //    && !Worksheet.Internals.RowsCollection[absoluteAddress.RowNumber].GetStyleId().Equals(worksheetStyleId))
                //    style = Worksheet.Internals.RowsCollection[absoluteAddress.RowNumber].Style;
                //else if (Worksheet.Internals.ColumnsCollection.ContainsKey(absoluteAddress.ColumnNumber)
                //         &&
                //         !Worksheet.Internals.ColumnsCollection[absoluteAddress.ColumnNumber].GetStyleId().Equals(worksheetStyleId))
                //    style = Worksheet.Internals.ColumnsCollection[absoluteAddress.ColumnNumber].Style;
            }
            var newCell = new XLCell(Worksheet, absoluteAddress, styleId);
            Worksheet.Internals.CellsCollection.Add(absoluteAddress.RowNumber, absoluteAddress.ColumnNumber, newCell);
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
            return Range(firstCell.Address, lastCell.Address);
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
            var newFirstCellAddress = (XLAddress)rangeAddress.FirstAddress + RangeAddress.FirstAddress - 1;
            newFirstCellAddress.FixedRow = rangeAddress.FirstAddress.FixedRow;
            newFirstCellAddress.FixedColumn = rangeAddress.FirstAddress.FixedColumn;

            var newLastCellAddress = (XLAddress)rangeAddress.LastAddress + RangeAddress.FirstAddress - 1;
            newLastCellAddress.FixedRow = rangeAddress.LastAddress.FixedRow;
            newLastCellAddress.FixedColumn = rangeAddress.LastAddress.FixedColumn;

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
                return ExcelHelper.GetColumnLetterFromNumber(test) + "1";
            return address;
        }

        public XLCells CellsUsed(bool includeFormats)
        {
            var cells = new XLCells(true, includeFormats) {RangeAddress};
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
            int columnCount = ColumnCount();
            int firstColumn = RangeAddress.FirstAddress.ColumnNumber + columnCount;
            if (firstColumn > ExcelHelper.MaxColumnNumber)
                firstColumn = ExcelHelper.MaxColumnNumber;
            int lastColumn = firstColumn + ColumnCount() - 1;
            if (lastColumn > ExcelHelper.MaxColumnNumber)
                lastColumn = ExcelHelper.MaxColumnNumber;

            int firstRow = RangeAddress.FirstAddress.RowNumber;
            int lastRow = firstRow + RowCount() - 1;
            if (lastRow > ExcelHelper.MaxRowNumber)
                lastRow = ExcelHelper.MaxRowNumber;

            var newRange = Worksheet.Range(firstRow, firstColumn, lastRow, lastColumn);
            return newRange.InsertColumnsBefore(onlyUsedCells, numberOfColumns, formatFromLeft);
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
            foreach (XLWorksheet ws in Worksheet.Workbook.WorksheetsInternal)
            {
                foreach (
                    XLCell cell in
                        ws.Internals.CellsCollection.GetCells(c => !StringExtensions.IsNullOrWhiteSpace(c.FormulaA1))
                    )
                    cell.ShiftFormulaColumns((XLRange)AsRange(), numberOfColumns);
            }

            var cellsToInsert = new Dictionary<IXLAddress, XLCell>();
            var cellsToDelete = new List<IXLAddress>();
            //var cellsToBlank = new List<IXLAddress>();
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
                            newCell.CopyValues(oldCell);
                            newCell.FormulaA1 = oldCell.FormulaA1;
                            cellsToInsert.Add(newKey, newCell);
                            cellsToDelete.Add(oldKey);
                            //if (oldKey.ColumnNumber < firstColumn + numberOfColumns)
                            //    cellsToBlank.Add(oldKey);
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
                    newCell.CopyValues(c);
                    newCell.FormulaA1 = c.FormulaA1;
                    cellsToInsert.Add(newKey, newCell);
                    cellsToDelete.Add(c.Address);
                    //if (c.Address.ColumnNumber < firstColumn + numberOfColumns)
                    //    cellsToBlank.Add(c.Address);
                }
            }
            cellsToDelete.ForEach(c => Worksheet.Internals.CellsCollection.Remove(c.RowNumber, c.ColumnNumber));
            cellsToInsert.ForEach(
                c => Worksheet.Internals.CellsCollection.Add(c.Key.RowNumber, c.Key.ColumnNumber, c.Value));

            Worksheet.NotifyRangeShiftedColumns((XLRange)AsRange(), numberOfColumns);
            var rangeToReturn = Worksheet.Range(
                RangeAddress.FirstAddress.RowNumber,
                RangeAddress.FirstAddress.ColumnNumber - numberOfColumns,
                RangeAddress.LastAddress.RowNumber,
                RangeAddress.LastAddress.ColumnNumber - numberOfColumns
                );

            if (formatFromLeft && rangeToReturn.RangeAddress.FirstAddress.ColumnNumber > 1)
            {
                var model = rangeToReturn.FirstColumn().ColumnLeft();
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
            int rowCount = RowCount();
            int firstRow = RangeAddress.FirstAddress.RowNumber + rowCount;
            if (firstRow > ExcelHelper.MaxRowNumber)
                firstRow = ExcelHelper.MaxRowNumber;
            int lastRow = firstRow + RowCount() - 1;
            if (lastRow > ExcelHelper.MaxRowNumber)
                lastRow = ExcelHelper.MaxRowNumber;

            int firstColumn = RangeAddress.FirstAddress.ColumnNumber;
            int lastColumn = firstColumn + ColumnCount() - 1;
            if (lastColumn > ExcelHelper.MaxColumnNumber)
                lastColumn = ExcelHelper.MaxColumnNumber;

            var newRange = Worksheet.Range(firstRow, firstColumn, lastRow, lastColumn);
            return newRange.InsertRowsAbove(onlyUsedCells, numberOfRows, formatFromAbove);
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

        public IXLRangeRows InsertRowsAbove(Boolean onlyUsedCells, Int32 numberOfRows, Boolean formatFromAbove = true)
        {
            foreach (XLWorksheet ws in Worksheet.Workbook.WorksheetsInternal)
            {
                foreach (
                    XLCell cell in
                        ws.Internals.CellsCollection.GetCells(c => !StringExtensions.IsNullOrWhiteSpace(c.FormulaA1))
                    )
                    cell.ShiftFormulaRows((XLRange)AsRange(), numberOfRows);
            }

            var cellsToInsert = new Dictionary<IXLAddress, XLCell>();
            var cellsToDelete = new List<IXLAddress>();
            //var cellsToBlank = new List<IXLAddress>();
            int firstRow = RangeAddress.FirstAddress.RowNumber;
            int firstColumn = RangeAddress.FirstAddress.ColumnNumber;
            int lastColumn = RangeAddress.FirstAddress.ColumnNumber + ColumnCount() - 1;

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
                            var oldCell = Worksheet.Internals.CellsCollection.GetCell(ro, co) ??
                                          Worksheet.Cell(oldKey);

                            var newCell = new XLCell(Worksheet, newKey, oldCell.GetStyleId());
                            newCell.CopyFrom(oldCell);
                            newCell.FormulaA1 = oldCell.FormulaA1;
                            cellsToInsert.Add(newKey, newCell);
                            cellsToDelete.Add(oldKey);
                            //if (oldKey.RowNumber < firstRow + numberOfRows)
                            //    cellsToBlank.Add(oldKey);
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
                    newCell.CopyFrom(c);
                    newCell.FormulaA1 = c.FormulaA1;
                    cellsToInsert.Add(newKey, newCell);
                    cellsToDelete.Add(c.Address);
                    //if (c.Address.RowNumber < firstRow + numberOfRows)
                    //    cellsToBlank.Add(c.Address);
                }
            }
            cellsToDelete.ForEach(c => Worksheet.Internals.CellsCollection.Remove(c.RowNumber, c.ColumnNumber));
            cellsToInsert.ForEach(
                c => Worksheet.Internals.CellsCollection.Add(c.Key.RowNumber, c.Key.ColumnNumber, c.Value));
            
            //foreach (IXLAddress c in cellsToBlank)
            //{
            //    IXLStyle styleToUse;
                
            //    styleToUse = Worksheet.Internals.ColumnsCollection.ContainsKey(c.ColumnNumber)
            //                         ? Worksheet.Internals.ColumnsCollection[c.ColumnNumber].Style
            //                         : Worksheet.Style;
            //    Worksheet.Cell(c.RowNumber, c.ColumnNumber).Style = styleToUse;
            //}
            Worksheet.NotifyRangeShiftedRows((XLRange)AsRange(), numberOfRows);
            var rangeToReturn = Worksheet.Range(
                RangeAddress.FirstAddress.RowNumber - numberOfRows,
                RangeAddress.FirstAddress.ColumnNumber,
                RangeAddress.LastAddress.RowNumber - numberOfRows,
                RangeAddress.LastAddress.ColumnNumber
                );

            if (formatFromAbove && rangeToReturn.RangeAddress.FirstAddress.RowNumber > 1)
            {
                var model = rangeToReturn.FirstRow().RowAbove();
                var modelFirstColumn = model.FirstCellUsed(true);
                var modelLastColumn = model.LastCellUsed(true);
                if (modelLastColumn != null)
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

            return rangeToReturn.Rows();
        }

        private void ClearMerged()
        {
            var mergeToDelete = Worksheet.Internals.MergedRanges.Where(Intersects).ToList();
            mergeToDelete.ForEach(m => Worksheet.Internals.MergedRanges.Remove(m));
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
                            c => !StringExtensions.IsNullOrWhiteSpace(c.FormulaA1))))
            {
                if (shiftDeleteCells == XLShiftDeletedCells.ShiftCellsUp)
                    cell.ShiftFormulaRows((XLRange)shiftedRangeFormula, numberOfRows * -1);
                else
                    cell.ShiftFormulaColumns((XLRange)shiftedRangeFormula, numberOfColumns * -1);
            }

            // Range to shift...
            var cellsToInsert = new Dictionary<IXLAddress, XLCell>();
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
                newCell.CopyValues(c);
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

            var shiftedRange = (XLRange)AsRange();
            if (shiftDeleteCells == XLShiftDeletedCells.ShiftCellsUp)
                Worksheet.NotifyRangeShiftedRows(shiftedRange, rowModifier * -1);
            else
                Worksheet.NotifyRangeShiftedColumns(shiftedRange, columnModifier * -1);
        }

        public override string ToString()
        {
            return String.Format("'{0}'!{1}:{2}", Worksheet.Name, RangeAddress.FirstAddress, RangeAddress.LastAddress);
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
                 &&
                 thisRangeAddress.FirstAddress.ColumnNumber >= shiftedRange.RangeAddress.FirstAddress.ColumnNumber
                 && thisRangeAddress.LastAddress.ColumnNumber <= shiftedRange.RangeAddress.FirstAddress.ColumnNumber
                 // all rows
                 && thisRangeAddress.FirstAddress.RowNumber >= shiftedRange.RangeAddress.FirstAddress.RowNumber
                 &&
                 thisRangeAddress.LastAddress.RowNumber <=
                 shiftedRange.RangeAddress.LastAddress.RowNumber - rowsShifted
                ) || (
                         shiftedRange.RangeAddress.FirstAddress.RowNumber <= thisRangeAddress.FirstAddress.RowNumber
                         &&
                         shiftedRange.RangeAddress.FirstAddress.ColumnNumber <=
                         thisRangeAddress.FirstAddress.ColumnNumber
                         &&
                         shiftedRange.RangeAddress.LastAddress.ColumnNumber >=
                         thisRangeAddress.LastAddress.ColumnNumber
                         && shiftedRange.RowCount() >
                         (thisRangeAddress.LastAddress.RowNumber - thisRangeAddress.FirstAddress.RowNumber + 1)
                         +
                         (thisRangeAddress.FirstAddress.RowNumber - shiftedRange.RangeAddress.FirstAddress.RowNumber)))
                thisRangeAddress.IsInvalid = true;
            else
            {
                if (shiftedRange.RangeAddress.FirstAddress.ColumnNumber <=
                    thisRangeAddress.FirstAddress.ColumnNumber
                    &&
                    shiftedRange.RangeAddress.LastAddress.ColumnNumber >= thisRangeAddress.LastAddress.ColumnNumber)
                {
                    if (
                        (shiftedRange.RangeAddress.FirstAddress.RowNumber <= thisRangeAddress.FirstAddress.RowNumber &&
                         rowsShifted > 0)
                        ||
                        (shiftedRange.RangeAddress.FirstAddress.RowNumber < thisRangeAddress.FirstAddress.RowNumber &&
                         rowsShifted < 0)
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

        public IXLRange RangeUsed(bool includeStyles)
        {
            var firstCell = FirstCellUsed(includeStyles);
            if (firstCell == null)
                return null;
            var lastCell = LastCellUsed(includeStyles);
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

        public void SetAutoFilter()
        {
            SetAutoFilter(true);
        }

        public void SetAutoFilter(Boolean autoFilter)
        {
            Worksheet.AutoFilterRange = autoFilter ? this : null;
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
        public XLPivotTable CreatePivotTable(IXLCell targetCell)
        {
            return CreatePivotTable(targetCell, Guid.NewGuid().ToString());
        }
        IXLPivotTable IXLRangeBase.CreatePivotTable(IXLCell targetCell, String name)
        {
            return CreatePivotTable(targetCell, name);
        }
        public XLPivotTable CreatePivotTable(IXLCell targetCell, String name)
        {
            return (XLPivotTable)this.Worksheet.PivotTables.AddNew(name, targetCell, this.AsRange());
        }
    }
}