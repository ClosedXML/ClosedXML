using ClosedXML.Extensions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{

    internal abstract class XLRangeBase : XLStylizedBase, IXLRangeBase, IXLStylized
    {
        #region Fields

        private XLSortElements _sortRows;
        private XLSortElements _sortColumns;

        #endregion Fields
        
        protected IXLStyle GetStyle()
        {
            return Style;
        }

        #region Constructor

        private static Int32 IdCounter = 0;
        private readonly Int32 Id;

        protected XLRangeBase(XLRangeAddress rangeAddress, XLStyleValue styleValue)
            : base(styleValue)
        {
            Id = ++IdCounter;

            _rangeAddress = rangeAddress;
        }

        #endregion Constructor

        protected virtual void OnRangeAddressChanged(XLRangeAddress oldAddress, XLRangeAddress newAddress)
        {
            Worksheet.RellocateRange(RangeType, oldAddress, newAddress);
        }

        #region Public properties

        private XLRangeAddress _rangeAddress;

        public XLRangeAddress RangeAddress
        {
            get { return _rangeAddress; }
            protected set
            {
                if (_rangeAddress != value)
                {
                    var oldAddress = _rangeAddress;
                    _rangeAddress = value;
                    OnRangeAddressChanged(oldAddress, _rangeAddress);
                }
            }
        }

        public XLWorksheet Worksheet
        {
            get { return RangeAddress.Worksheet; }
        }

        public IXLDataValidation NewDataValidation
        {
            get
            {
                var newRanges = new XLRanges { AsRange() };
                var dataValidation = DataValidation;

                if (dataValidation != null)
                    Worksheet.DataValidations.Delete(dataValidation);

                dataValidation = new XLDataValidation(newRanges);
                Worksheet.DataValidations.Add(dataValidation);
                return dataValidation;
            }
        }

        /// <summary>
        /// Get the data validation rule containing current range or create a new one if no rule was defined for range.
        /// </summary>
        public IXLDataValidation DataValidation
        {
            get
            {
                return SetDataValidation();
            }
        }

        private IXLDataValidation GetDataValidation()
        {
            foreach (var xlDataValidation in Worksheet.DataValidations)
            {
                foreach (var range in xlDataValidation.Ranges)
                {
                    if (range.ToString() == ToString())
                        return xlDataValidation;
                }
            }
            return null;
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

        public XLDataType DataType
        {
            set { Cells().ForEach(c => c.DataType = value); }
        }

        #endregion IXLRangeBase Members

        #region IXLStylized Members

        public override IXLRanges RangesUsed
        {
            get
            {
                var retVal = new XLRanges { AsRange() };
                return retVal;
            }
        }

        protected override IEnumerable<XLStylizedBase> Children
        {
            get
            {
                foreach (var cell in Cells().OfType<XLCell>())
                    yield return cell;
            }
        }

        public override IEnumerable<IXLStyle> Styles
        {
            get
            {
                foreach (IXLCell cell in Cells())
                    yield return cell.Style;
            }
        }
        #endregion IXLStylized Members

        #endregion Public properties

        #region IXLRangeBase Members

        IXLCell IXLRangeBase.FirstCell()
        {
            return FirstCell();
        }

        IXLCell IXLRangeBase.LastCell()
        {
            return LastCell();
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLCell IXLRangeBase.FirstCellUsed()
        {
            return FirstCellUsed(false);
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLCell IXLRangeBase.FirstCellUsed(bool includeFormats)
        {
            return FirstCellUsed(includeFormats);
        }

        IXLCell IXLRangeBase.FirstCellUsed(XLCellsUsedOptions options)
        {
            return FirstCellUsed(options, null);
        }

        IXLCell IXLRangeBase.FirstCellUsed(Func<IXLCell, Boolean> predicate)
        {
            return FirstCellUsed(predicate);
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLCell IXLRangeBase.FirstCellUsed(Boolean includeFormats, Func<IXLCell, Boolean> predicate)
        {
            return FirstCellUsed(includeFormats, predicate);
        }

        IXLCell IXLRangeBase.FirstCellUsed(XLCellsUsedOptions options, Func<IXLCell, Boolean> predicate)
        {
            return FirstCellUsed(options, predicate);
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLCell IXLRangeBase.LastCellUsed()
        {
            return LastCellUsed(false);
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLCell IXLRangeBase.LastCellUsed(bool includeFormats)
        {
            return LastCellUsed(includeFormats);
        }

        IXLCell IXLRangeBase.LastCellUsed(XLCellsUsedOptions options)
        {
            return LastCellUsed(options, null);
        }

        IXLCell IXLRangeBase.LastCellUsed(Func<IXLCell, Boolean> predicate)
        {
            return LastCellUsed(predicate);
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLCell IXLRangeBase.LastCellUsed(Boolean includeFormats, Func<IXLCell, Boolean> predicate)
        {
            return LastCellUsed(includeFormats, predicate);
        }

        IXLCell IXLRangeBase.LastCellUsed(XLCellsUsedOptions options, Func<IXLCell, Boolean> predicate)
        {
            return LastCellUsed(options, predicate);
        }

        public IXLCells Cells()
        {
            return Cells(false);
        }

        public IXLCells Cells(Boolean usedCellsOnly)
        {
            return Cells(usedCellsOnly, XLCellsUsedOptions.AllContents);
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        public IXLCells Cells(Boolean usedCellsOnly, Boolean includeFormats)
        {
            return Cells(usedCellsOnly, includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents
            );
        }

        public IXLCells Cells(Boolean usedCellsOnly, XLCellsUsedOptions options)
        {
            var cells = new XLCells(usedCellsOnly, options) { RangeAddress };
            return cells;
        }


        public IXLCells Cells(String cells)
        {
            return Ranges(cells).Cells();
        }

        public IXLCells Cells(Func<IXLCell, Boolean> predicate)
        {
            var cells = new XLCells(false, XLCellsUsedOptions.AllContents, predicate) { RangeAddress };
            return cells;
        }

        public IXLCells CellsUsed()
        {
            return Cells(true);
        }

        /// <summary>
        /// Return the collection of cell values not initializing empty cells.
        /// </summary>
        public IEnumerable CellValues()
        {
            for (int ro = RangeAddress.FirstAddress.RowNumber; ro <= RangeAddress.LastAddress.RowNumber; ro++)
            {
                for (int co = RangeAddress.FirstAddress.ColumnNumber; co <= RangeAddress.LastAddress.ColumnNumber; co++)
                {
                    yield return Worksheet.GetCellValue(ro, co);
                }
            }
        }

        public IXLRange Merge()
        {
            return Merge(true);
        }

        public IXLRange Merge(Boolean checkIntersect)
        {
            if (checkIntersect)
            {
                var intersectedMergedRanges = Worksheet.Internals.MergedRanges.GetIntersectedRanges(RangeAddress).ToList();
                foreach (var intersectedMergedRange in intersectedMergedRanges)
                {
                    Worksheet.Internals.MergedRanges.Remove(intersectedMergedRange);
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

        public IXLRangeBase Clear(XLClearOptions clearOptions = XLClearOptions.All)
        {
            var options = clearOptions.ToCellsUsedOptions();
            foreach (var cell in CellsUsed(options))
            {
                // We'll clear the conditional formatting later down.
                (cell as XLCell).Clear(clearOptions & ~XLClearOptions.ConditionalFormats, true);
            }

            if (clearOptions.HasFlag(XLClearOptions.NormalFormats))
                ClearMerged();

            if (clearOptions.HasFlag(XLClearOptions.ConditionalFormats))
                RemoveConditionalFormatting();

            if (clearOptions == XLClearOptions.All)
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

        internal void RemoveConditionalFormatting()
        {
            var mf = RangeAddress.FirstAddress;
            var ml = RangeAddress.LastAddress;
            foreach (var format in Worksheet.ConditionalFormats.Where(x => x.Ranges.GetIntersectedRanges(RangeAddress).Any()).ToList())
            {
                var cfRanges = format.Ranges.ToList();
                format.Ranges.RemoveAll();

                foreach (var cfRange in cfRanges)
                {
                    if (!cfRange.Intersects(this))
                    {
                        format.Ranges.Add(cfRange);
                        continue;
                    }

                    var f = cfRange.RangeAddress.FirstAddress;
                    var l = cfRange.RangeAddress.LastAddress;
                    bool byWidth = false, byHeight = false;
                    XLRange rng1 = null, rng2 = null;
                    if (mf.ColumnNumber <= f.ColumnNumber && ml.ColumnNumber >= l.ColumnNumber)
                    {
                        if (mf.RowNumber.Between(f.RowNumber, l.RowNumber) || ml.RowNumber.Between(f.RowNumber, l.RowNumber))
                        {
                            if (mf.RowNumber > f.RowNumber)
                                rng1 = Worksheet.Range(f.RowNumber, f.ColumnNumber, mf.RowNumber - 1, l.ColumnNumber);
                            if (ml.RowNumber < l.RowNumber)
                                rng2 = Worksheet.Range(ml.RowNumber + 1, f.ColumnNumber, l.RowNumber, l.ColumnNumber);
                        }
                        byWidth = true;
                    }

                    if (mf.RowNumber <= f.RowNumber && ml.RowNumber >= l.RowNumber)
                    {
                        if (mf.ColumnNumber.Between(f.ColumnNumber, l.ColumnNumber) || ml.ColumnNumber.Between(f.ColumnNumber, l.ColumnNumber))
                        {
                            if (mf.ColumnNumber > f.ColumnNumber)
                                rng1 = Worksheet.Range(f.RowNumber, f.ColumnNumber, l.RowNumber, mf.ColumnNumber - 1);
                            if (ml.ColumnNumber < l.ColumnNumber)
                                rng2 = Worksheet.Range(f.RowNumber, ml.ColumnNumber + 1, l.RowNumber, l.ColumnNumber);
                        }
                        byHeight = true;
                    }

                    if (rng1 != null)
                    {
                        format.Ranges.Add(rng1);
                    }
                    if (rng2 != null)
                    {
                        //TODO: reflect the formula for a new range
                        format.Ranges.Add(rng2);
                    }

                    if (!byWidth && !byHeight)
                        format.Ranges.Add(cfRange); // Not split, preserve original
                }
                if (!format.Ranges.Any())
                    Worksheet.ConditionalFormats.Remove(x => x == format);
            }
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
            return Intersects(Worksheet.Range(rangeAddress));
        }

        public bool Intersects(IXLRangeBase range)
        {
            if (!range.RangeAddress.IsValid || !RangeAddress.IsValid)
                return false;
            var ma = range.RangeAddress;
            var ra = RangeAddress;
            return ra.Intersects(ma);
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

        public virtual Boolean IsEmpty()
        {
            return !CellsUsed().Any() || CellsUsed().Any(c => c.IsEmpty());
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        public virtual Boolean IsEmpty(Boolean includeFormats)
        {
            return IsEmpty(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents);
        }

        public virtual Boolean IsEmpty(XLCellsUsedOptions options)
        {
            return CellsUsed(options).Cast<XLCell>().All(c => c.IsEmpty(options));
        }

        public virtual Boolean IsEntireRow()
        {
            return RangeAddress.FirstAddress.ColumnNumber == 1
                   && RangeAddress.LastAddress.ColumnNumber == XLHelper.MaxColumnNumber;
        }

        public virtual Boolean IsEntireColumn()
        {
            return RangeAddress.FirstAddress.RowNumber == 1
                   && RangeAddress.LastAddress.RowNumber == XLHelper.MaxRowNumber;
        }

        #endregion IXLRangeBase Members
        
        public IXLCells Search(String searchText, CompareOptions compareOptions = CompareOptions.Ordinal, Boolean searchFormulae = false)
        {
            var culture = CultureInfo.CurrentCulture;
            return CellsUsed(XLCellsUsedOptions.AllContents, c =>
            {
                try
                {
                    if (searchFormulae)
                        return c.HasFormula
                               && culture.CompareInfo.IndexOf(c.FormulaA1, searchText, compareOptions) >= 0
                               || culture.CompareInfo.IndexOf(c.Value.ToString(), searchText, compareOptions) >= 0;
                    else
                        return culture.CompareInfo.IndexOf(c.GetFormattedString(), searchText, compareOptions) >= 0;
                }
                catch
                {
                    return false;
                }
            });
        }

        internal XLCell FirstCell()
        {
            return Cell(1, 1);
        }

        internal XLCell LastCell()
        {
            return Cell(RowCount(), ColumnCount());
        }

        internal XLCell FirstCellUsed()
        {
            return FirstCellUsed(XLCellsUsedOptions.AllContents, predicate: null);
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        internal XLCell FirstCellUsed(Boolean includeFormats)
        {
            return FirstCellUsed(includeFormats, null);
        }


        internal XLCell FirstCellUsed(Func<IXLCell, Boolean> predicate)
        {
            return FirstCellUsed(XLCellsUsedOptions.AllContents, predicate);
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        internal XLCell FirstCellUsed(Boolean includeFormats, Func<IXLCell, Boolean> predicate)
        {
            return FirstCellUsed(includeFormats
                    ? XLCellsUsedOptions.All
                    : XLCellsUsedOptions.AllContents,
                predicate);
        }

        internal XLCell FirstCellUsed(XLCellsUsedOptions options, Func<IXLCell, Boolean> predicate)
        {
            var cellsUsed = CellsUsed(options, predicate).ToList();

            if (!cellsUsed.Any())
                return null;

            var firstRow = cellsUsed.Min(c => c.Address.RowNumber);
            var firstColumn = cellsUsed.Min(c => c.Address.ColumnNumber);

            return Worksheet.Cell(firstRow, firstColumn);
        }

        internal XLCell LastCellUsed()
        {
            return LastCellUsed(XLCellsUsedOptions.AllContents, predicate: null);
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        internal XLCell LastCellUsed(Boolean includeFormats)
        {
            return LastCellUsed(includeFormats, null);
        }

        internal XLCell LastCellUsed(Func<IXLCell, Boolean> predicate)
        {
            return LastCellUsed(XLCellsUsedOptions.AllContents, predicate);
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        internal XLCell LastCellUsed(Boolean includeFormats, Func<IXLCell, Boolean> predicate)
        {
            return LastCellUsed(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents,
                predicate);
        }

        internal XLCell LastCellUsed(XLCellsUsedOptions options, Func<IXLCell, Boolean> predicate)
        {
            var cellsUsed = CellsUsed(options, predicate).ToList();

            if (!cellsUsed.Any())
                return null;

            var lastRow = cellsUsed.Max(c => c.Address.RowNumber);
            var lastColumn = cellsUsed.Max(c => c.Address.ColumnNumber);

            return Worksheet.Cell(lastRow, lastColumn);
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

        public XLCell Cell(in XLAddress cellAddressInRange)
        {
            Int32 absRow = cellAddressInRange.RowNumber + RangeAddress.FirstAddress.RowNumber - 1;
            Int32 absColumn = cellAddressInRange.ColumnNumber + RangeAddress.FirstAddress.ColumnNumber - 1;

            if (absRow <= 0 || absRow > XLHelper.MaxRowNumber)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(cellAddressInRange),
                    String.Format("Row number must be between 1 and {0}", XLHelper.MaxRowNumber)
                );
            }

            if (absColumn <= 0 || absColumn > XLHelper.MaxColumnNumber)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(cellAddressInRange),
                    String.Format("Column number must be between 1 and {0}", XLHelper.MaxColumnNumber)
                );
            }

            var cell = Worksheet.Internals.CellsCollection.GetCell(absRow,
                                                                   absColumn);

            if (cell != null)
                return cell;

            var styleValue = this.StyleValue;

            if (styleValue == Worksheet.StyleValue)
            {
                if (Worksheet.Internals.RowsCollection.TryGetValue(absRow, out XLRow row)
                    && row.StyleValue != Worksheet.StyleValue)
                    styleValue = row.StyleValue;
                else if (Worksheet.Internals.ColumnsCollection.TryGetValue(absColumn, out XLColumn column)
                    && column.StyleValue != Worksheet.StyleValue)
                    styleValue = column.StyleValue;
            }
            var absoluteAddress = new XLAddress(this.Worksheet,
                                 absRow,
                                 absColumn,
                                 cellAddressInRange.FixedRow,
                                 cellAddressInRange.FixedColumn);

            // If the default style for this range base is empty, but the worksheet
            // has a default style, use the worksheet's default style
            XLCell newCell = new XLCell(Worksheet, absoluteAddress, styleValue);

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

        internal abstract void WorksheetRangeShiftedColumns(XLRange range, int columnsShifted);

        internal abstract void WorksheetRangeShiftedRows(XLRange range, int rowsShifted);

        public abstract XLRangeType RangeType { get; }

        public XLRange Range(IXLCell firstCell, IXLCell lastCell)
        {
            var newFirstCellAddress = (XLAddress)firstCell.Address;
            var newLastCellAddress = (XLAddress)lastCell.Address;

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

            if (newFirstCellAddress.Worksheet != null)
                return newFirstCellAddress.Worksheet.GetOrCreateRange(xlRangeParameters);
            else if (Worksheet != null)
                return Worksheet.GetOrCreateRange(xlRangeParameters);
            else
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
            var rangeAddress = new XLRangeAddress((XLAddress)firstCellAddress, (XLAddress)lastCellAddress);
            return Range(rangeAddress);
        }

        public XLRange Range(IXLRangeAddress rangeAddress)
        {
            var ws = (XLWorksheet) rangeAddress.FirstAddress.Worksheet ??
                     (XLWorksheet) rangeAddress.LastAddress.Worksheet ??
                     Worksheet;
            var newFirstCellAddress = new XLAddress(ws,
                                 rangeAddress.FirstAddress.RowNumber + RangeAddress.FirstAddress.RowNumber - 1,
                                 rangeAddress.FirstAddress.ColumnNumber + RangeAddress.FirstAddress.ColumnNumber - 1,
                                 rangeAddress.FirstAddress.FixedRow,
                                 rangeAddress.FirstAddress.FixedColumn);

            var newLastCellAddress = new XLAddress(ws,
                                rangeAddress.LastAddress.RowNumber + RangeAddress.FirstAddress.RowNumber - 1,
                                rangeAddress.LastAddress.ColumnNumber + RangeAddress.FirstAddress.ColumnNumber - 1,
                                rangeAddress.LastAddress.FixedRow,
                                rangeAddress.LastAddress.FixedColumn);

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
            if (Int32.TryParse(address, out Int32 test))
                return "A" + address;
            return address;
        }

        protected String FixRowAddress(String address)
        {
            if (Int32.TryParse(address, out Int32 test))
                return XLHelper.GetColumnLetterFromNumber(test) + "1";
            return address;
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        public IXLCells CellsUsed(bool includeFormats)
        {
            return CellsUsed(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents);
        }

        public IXLCells CellsUsed(XLCellsUsedOptions options)
        {
            var cells = new XLCells(true, options) { RangeAddress };
            return cells;
        }

        public IXLCells CellsUsed(Func<IXLCell, Boolean> predicate)
        {
            var cells = new XLCells(true, XLCellsUsedOptions.AllContents, predicate) { RangeAddress };
            return cells;
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        public IXLCells CellsUsed(Boolean includeFormats, Func<IXLCell, Boolean> predicate)
        {
            return CellsUsed(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents,
                predicate);
        }

        public IXLCells CellsUsed(XLCellsUsedOptions options, Func<IXLCell, Boolean> predicate)
        {
            var cells = new XLCells(true, options, predicate) { RangeAddress };
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
            if (numberOfColumns <= 0 || numberOfColumns > XLHelper.MaxColumnNumber)
                throw new ArgumentOutOfRangeException(nameof(numberOfColumns),
                    $"Number of columns to insert must be a positive number no more than {XLHelper.MaxColumnNumber}");

            foreach (XLWorksheet ws in Worksheet.Workbook.WorksheetsInternal)
            {
                foreach (XLCell cell in ws.Internals.CellsCollection.GetCells(c => !String.IsNullOrWhiteSpace(c.FormulaA1)))
                    cell.ShiftFormulaColumns(AsRange(), numberOfColumns);
            }

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
                        int newColumn = co + numberOfColumns;
                        for (int ro = lastRow; ro >= firstRow; ro--)
                        {
                            var oldKey = new XLAddress(Worksheet, ro, co, false, false);
                            var newKey = new XLAddress(Worksheet, ro, newColumn, false, false);
                            var oldCell = Worksheet.Internals.CellsCollection.GetCell(ro, co) ??
                                          Worksheet.Cell(oldKey);

                            var newCell = new XLCell(Worksheet, newKey, oldCell.StyleValue);
                            newCell.CopyValuesFrom(oldCell);
                            newCell.FormulaA1 = oldCell.FormulaA1;
                            cellsToInsert.Add(newKey, newCell);
                            cellsToDelete.Add(oldKey);
                        }

                        if (this.IsEntireColumn())
                        {
                            Worksheet.Column(newColumn).Width = Worksheet.Column(co).Width;
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
                    var newCell = new XLCell(Worksheet, newKey, c.StyleValue);
                    newCell.CopyValuesFrom(c);
                    newCell.FormulaA1 = c.FormulaA1;
                    cellsToInsert.Add(newKey, newCell);
                    cellsToDelete.Add(c.Address);
                }
            }

            cellsToDelete.ForEach(c => Worksheet.Internals.CellsCollection.Remove(c.RowNumber, c.ColumnNumber));
            cellsToInsert.ForEach(
                c => Worksheet.Internals.CellsCollection.Add(c.Key.RowNumber, c.Key.ColumnNumber, c.Value));

            Int32 firstRowReturn = RangeAddress.FirstAddress.RowNumber;
            Int32 lastRowReturn = RangeAddress.LastAddress.RowNumber;
            Int32 firstColumnReturn = RangeAddress.FirstAddress.ColumnNumber;
            Int32 lastColumnReturn = RangeAddress.FirstAddress.ColumnNumber + numberOfColumns - 1;

            Worksheet.NotifyRangeShiftedColumns(AsRange(), numberOfColumns);

            var rangeToReturn = Worksheet.Range(firstRowReturn, firstColumnReturn, lastRowReturn, lastColumnReturn);

            // We deliberately ignore conditional formats and data validation here. Their shifting is handled elsewhere
            var contentFlags = XLCellsUsedOptions.All
                & ~XLCellsUsedOptions.ConditionalFormats
                & ~XLCellsUsedOptions.DataValidation;

            if (formatFromLeft && rangeToReturn.RangeAddress.FirstAddress.ColumnNumber > 1)
            {
                var firstColumnUsed = rangeToReturn.FirstColumn();
                var model = firstColumnUsed.ColumnLeft();
                        var modelFirstRow = (model as IXLRangeBase).FirstCellUsed(contentFlags);
                        var modelLastRow = (model as IXLRangeBase).LastCellUsed(contentFlags);
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
                var lastRoUsed = rangeToReturn.LastRowUsed(contentFlags);
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
            if (numberOfRows <= 0 || numberOfRows > XLHelper.MaxRowNumber)
                throw new ArgumentOutOfRangeException(nameof(numberOfRows),
                    $"Number of rows to insert must be a positive number no more than {XLHelper.MaxRowNumber}");

            var asRange = AsRange();
            foreach (XLWorksheet ws in Worksheet.Workbook.WorksheetsInternal)
            {
                foreach (XLCell cell in ws.Internals.CellsCollection.GetCells(c => !String.IsNullOrWhiteSpace(c.FormulaA1)))
                    cell.ShiftFormulaRows(asRange, numberOfRows);
            }

            var cellsToInsert = new Dictionary<IXLAddress, XLCell>();
            var cellsToDelete = new List<IXLAddress>();
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
                        int newRow = ro + numberOfRows;

                        for (int co = lastColumn; co >= firstColumn; co--)
                        {
                            var oldKey = new XLAddress(Worksheet, ro, co, false, false);
                            var newKey = new XLAddress(Worksheet, newRow, co, false, false);
                            var oldCell = Worksheet.Internals.CellsCollection.GetCell(ro, co);
                            if (oldCell != null)
                            {
                                var newCell = new XLCell(Worksheet, newKey, oldCell.StyleValue);
                                newCell.CopyValuesFrom(oldCell);
                                newCell.FormulaA1 = oldCell.FormulaA1;
                                cellsToInsert.Add(newKey, newCell);
                                cellsToDelete.Add(oldKey);
                            }
                        }
                        if (this.IsEntireRow())
                        {
                            Worksheet.Row(newRow).Height = Worksheet.Row(ro).Height;
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
                    var newCell = new XLCell(Worksheet, newKey, c.StyleValue);
                    newCell.CopyValuesFrom(c);
                    newCell.FormulaA1 = c.FormulaA1;
                    cellsToInsert.Add(newKey, newCell);
                    cellsToDelete.Add(c.Address);
                }
            }

            cellsToDelete.ForEach(c => Worksheet.Internals.CellsCollection.Remove(c.RowNumber, c.ColumnNumber));
            cellsToInsert.ForEach(c => Worksheet.Internals.CellsCollection.Add(c.Key.RowNumber, c.Key.ColumnNumber, c.Value));

            Int32 firstRowReturn = RangeAddress.FirstAddress.RowNumber;
            Int32 lastRowReturn = RangeAddress.FirstAddress.RowNumber + numberOfRows - 1;
            Int32 firstColumnReturn = RangeAddress.FirstAddress.ColumnNumber;
            Int32 lastColumnReturn = RangeAddress.LastAddress.ColumnNumber;

            Worksheet.NotifyRangeShiftedRows(AsRange(), numberOfRows);

            var rangeToReturn = Worksheet.Range(firstRowReturn, firstColumnReturn, lastRowReturn, lastColumnReturn);

            // We deliberately ignore conditional formats and data validation here. Their shifting is handled elsewhere
            var contentFlags = XLCellsUsedOptions.All
                & ~XLCellsUsedOptions.ConditionalFormats
                & ~XLCellsUsedOptions.DataValidation;

            if (formatFromAbove && rangeToReturn.RangeAddress.FirstAddress.RowNumber > 1)
            {
                var fr = rangeToReturn.FirstRow();
                var model = fr.RowAbove();
                var modelFirstColumn = (model as IXLRangeBase).FirstCellUsed(contentFlags);
                var modelLastColumn = (model as IXLRangeBase).LastCellUsed(contentFlags);
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
            else
            {
                var lastCoUsed = rangeToReturn.LastColumnUsed(contentFlags);
                if (lastCoUsed != null)
                {
                    Int32 lastCoReturned = lastCoUsed.ColumnNumber();
                    for (Int32 co = 1; co <= lastCoReturned; co++)
                    {
                        var styleToUse = Worksheet.Internals.ColumnsCollection.ContainsKey(co)
                                             ? Worksheet.Internals.ColumnsCollection[co].Style
                                             : Worksheet.Style;
                        rangeToReturn.Style = styleToUse;
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
            var mergeToDelete = Worksheet.Internals.MergedRanges.GetIntersectedRanges(RangeAddress).ToList();
            mergeToDelete.ForEach(m => Worksheet.Internals.MergedRanges.Remove(m));
        }

        public Boolean Contains(IXLCell cell)
        {
            return Contains((XLAddress)cell.Address);
        }

        public bool Contains(XLAddress first, XLAddress last)
        {
            return Contains(first) && Contains(last);
        }

        public bool Contains(XLAddress address)
        {
            return RangeAddress.Contains(in address);
        }

        public void Delete(XLShiftDeletedCells shiftDeleteCells)
        {
            int numberOfRows = RowCount();
            int numberOfColumns = ColumnCount();

            if (!RangeAddress.IsValid) return;

            IXLRange shiftedRangeFormula = Worksheet.Range(
                RangeAddress.FirstAddress.RowNumber,
                RangeAddress.FirstAddress.ColumnNumber,
                RangeAddress.LastAddress.RowNumber,
                RangeAddress.LastAddress.ColumnNumber);

            foreach (
                XLCell cell in
                    Worksheet.Workbook.Worksheets.Cast<XLWorksheet>().SelectMany(
                        xlWorksheet => (xlWorksheet).Internals.CellsCollection.GetCells(
                            c => !String.IsNullOrWhiteSpace(c.FormulaA1))))
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
                var newCell = new XLCell(Worksheet, newKey, c.StyleValue);
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

            var shiftedRange = AsRange();
            if (shiftDeleteCells == XLShiftDeletedCells.ShiftCellsUp)
                Worksheet.NotifyRangeShiftedRows(shiftedRange, rowModifier * -1);
            else
                Worksheet.NotifyRangeShiftedColumns(shiftedRange, columnModifier * -1);

            Worksheet.DeleteRange(RangeAddress);
        }

        public override string ToString()
        {
            return String.Concat(
                Worksheet.Name.EscapeSheetName(),
                '!',
                RangeAddress.FirstAddress,
                ':',
                RangeAddress.LastAddress);
        }

        protected IXLRangeAddress ShiftColumns(IXLRangeAddress thisRangeAddress, XLRange shiftedRange, int columnsShifted)
        {
            if (!thisRangeAddress.IsValid || !shiftedRange.RangeAddress.IsValid) return thisRangeAddress;

            bool allRowsAreCovered = thisRangeAddress.FirstAddress.RowNumber >= shiftedRange.RangeAddress.FirstAddress.RowNumber &&
                                     thisRangeAddress.LastAddress.RowNumber <= shiftedRange.RangeAddress.LastAddress.RowNumber;

            if (!allRowsAreCovered)
                return thisRangeAddress;

            bool shiftLeftBoundary = (columnsShifted > 0 && thisRangeAddress.FirstAddress.ColumnNumber >= shiftedRange.RangeAddress.FirstAddress.ColumnNumber) ||
                                     (columnsShifted < 0 && thisRangeAddress.FirstAddress.ColumnNumber > shiftedRange.RangeAddress.FirstAddress.ColumnNumber);

            bool shiftRightBoundary = thisRangeAddress.LastAddress.ColumnNumber >= shiftedRange.RangeAddress.FirstAddress.ColumnNumber;

            int newLeftBoundary = thisRangeAddress.FirstAddress.ColumnNumber;
            if (shiftLeftBoundary)
            {
                if (newLeftBoundary + columnsShifted > shiftedRange.RangeAddress.FirstAddress.ColumnNumber)
                    newLeftBoundary = newLeftBoundary + columnsShifted;
                else
                    newLeftBoundary = shiftedRange.RangeAddress.FirstAddress.ColumnNumber;
            }

            int newRightBoundary = thisRangeAddress.LastAddress.ColumnNumber;
            if (shiftRightBoundary)
                newRightBoundary += columnsShifted;

            bool destroyedByShift = newRightBoundary < newLeftBoundary;

            var firstAddress = (XLAddress)thisRangeAddress.FirstAddress;
            var lastAddress =  (XLAddress)thisRangeAddress.LastAddress;

            if (destroyedByShift)
            {
                firstAddress = Worksheet.InvalidAddress;
                lastAddress = Worksheet.InvalidAddress;
                Worksheet.DeleteRange(RangeAddress);
            }

            if (shiftLeftBoundary)
                firstAddress = new XLAddress(Worksheet,
                                             thisRangeAddress.FirstAddress.RowNumber,
                                             newLeftBoundary,
                                             thisRangeAddress.FirstAddress.FixedRow,
                                             thisRangeAddress.FirstAddress.FixedColumn);

            if (shiftRightBoundary)
                lastAddress = new XLAddress(Worksheet,
                                            thisRangeAddress.LastAddress.RowNumber,
                                            newRightBoundary,
                                            thisRangeAddress.LastAddress.FixedRow,
                                            thisRangeAddress.LastAddress.FixedColumn);

            return new XLRangeAddress(firstAddress, lastAddress);
        }

        protected IXLRangeAddress ShiftRows(IXLRangeAddress thisRangeAddress, XLRange shiftedRange, int rowsShifted)
        {
            if (!thisRangeAddress.IsValid || !shiftedRange.RangeAddress.IsValid) return thisRangeAddress;

            bool allColumnsAreCovered = thisRangeAddress.FirstAddress.ColumnNumber >= shiftedRange.RangeAddress.FirstAddress.ColumnNumber &&
                                        thisRangeAddress.LastAddress.ColumnNumber <= shiftedRange.RangeAddress.LastAddress.ColumnNumber;

            if (!allColumnsAreCovered)
                return thisRangeAddress;

            bool shiftTopBoundary = (rowsShifted > 0 && thisRangeAddress.FirstAddress.RowNumber >= shiftedRange.RangeAddress.FirstAddress.RowNumber) ||
                                    (rowsShifted < 0 && thisRangeAddress.FirstAddress.RowNumber > shiftedRange.RangeAddress.FirstAddress.RowNumber);

            bool shiftBottomBoundary = thisRangeAddress.LastAddress.RowNumber >= shiftedRange.RangeAddress.FirstAddress.RowNumber;

            int newTopBoundary = thisRangeAddress.FirstAddress.RowNumber;
            if (shiftTopBoundary)
            {
                if (newTopBoundary + rowsShifted > shiftedRange.RangeAddress.FirstAddress.RowNumber)
                    newTopBoundary = newTopBoundary + rowsShifted;
                else
                    newTopBoundary = shiftedRange.RangeAddress.FirstAddress.RowNumber;
            }

            int newBottomBoundary = thisRangeAddress.LastAddress.RowNumber;
            if (shiftBottomBoundary)
                newBottomBoundary += rowsShifted;

            bool destroyedByShift = newBottomBoundary < newTopBoundary;

            var firstAddress = (XLAddress)thisRangeAddress.FirstAddress;
            var lastAddress = (XLAddress)thisRangeAddress.LastAddress;

            if (destroyedByShift)
            {
                firstAddress = Worksheet.InvalidAddress;
                lastAddress = Worksheet.InvalidAddress;
                Worksheet.DeleteRange(RangeAddress);
            }

            if (shiftTopBoundary)
                firstAddress = new XLAddress(Worksheet,
                                             newTopBoundary,
                                             thisRangeAddress.FirstAddress.ColumnNumber,
                                             thisRangeAddress.FirstAddress.FixedRow,
                                             thisRangeAddress.FirstAddress.FixedColumn);

            if (shiftBottomBoundary)
                lastAddress = new XLAddress(Worksheet,
                                            newBottomBoundary,
                                            thisRangeAddress.LastAddress.ColumnNumber,
                                            thisRangeAddress.LastAddress.FixedRow,
                                            thisRangeAddress.LastAddress.FixedColumn);

            return new XLRangeAddress(firstAddress, lastAddress);
        }

        public IXLRange RangeUsed()
        {
            return RangeUsed(XLCellsUsedOptions.AllContents);
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        public IXLRange RangeUsed(bool includeFormats)
        {
            return RangeUsed(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents);
        }

        public IXLRange RangeUsed(XLCellsUsedOptions options)
        {
            var firstCell = (this as IXLRangeBase).FirstCellUsed(options);
            if (firstCell == null)
                return null;
            var lastCell = (this as IXLRangeBase).LastCellUsed(options);
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

        IXLPivotTable IXLRangeBase.CreatePivotTable(IXLCell targetCell, String name)
        {
            return CreatePivotTable(targetCell, name);
        }

        public XLPivotTable CreatePivotTable(IXLCell targetCell, String name)
        {
            return (XLPivotTable)targetCell.Worksheet.PivotTables.Add(name, targetCell, AsRange());
        }

        public IXLAutoFilter SetAutoFilter()
        {
            return SetAutoFilter(true);
        }

        public IXLAutoFilter SetAutoFilter(Boolean value)
        {
            if (value)
                return Worksheet.AutoFilter.Set(this);
            else
                return Worksheet.AutoFilter.Clear();
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

        private String DefaultSortString()
        {
            var sb = new StringBuilder();
            Int32 maxColumn = ColumnCount();
            if (maxColumn == XLHelper.MaxColumnNumber)
                maxColumn = (this as IXLRangeBase).LastCellUsed(XLCellsUsedOptions.All).Address.ColumnNumber;
            for (int i = 1; i <= maxColumn; i++)
            {
                if (sb.Length > 0)
                    sb.Append(',');

                sb.Append(i);
            }

            return sb.ToString();
        }

        public IXLRangeBase Sort()
        {
            if (!SortColumns.Any())
            {
                return Sort(DefaultSortString());
            }

            SortRangeRows();
            return this;
        }

        public IXLRangeBase Sort(String columnsToSortBy, XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true)
        {
            SortColumns.Clear();
            if (String.IsNullOrWhiteSpace(columnsToSortBy))
            {
                columnsToSortBy = DefaultSortString();
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

                if (!Int32.TryParse(coString, out Int32 co))
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
                maxColumn = (this as IXLRangeBase).LastCellUsed(XLCellsUsedOptions.All).Address.ColumnNumber;

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
                maxRow = (this as IXLRangeBase).LastCellUsed(XLCellsUsedOptions.All).Address.RowNumber;

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
            if (beg == end)
                return;
            int pivot = SortRangeRows(beg, end);
            if (pivot > beg)
                SortingRangeRows(beg, pivot - 1);
            if (pivot < end)
                SortingRangeRows(pivot + 1, end);
        }

        #endregion Sort Rows

        #region Sort Columns

        private void SortRangeColumns()
        {
            Int32 maxColumn = ColumnCount();
            if (maxColumn == XLHelper.MaxColumnNumber)
                maxColumn = (this as IXLRangeBase).LastCellUsed(XLCellsUsedOptions.All).Address.ColumnNumber;
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

        #endregion Sort Columns

        #endregion Sort

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
            return Worksheet.RangeColumn(new XLRangeAddress(firstCellAddress, lastCellAddress));
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

            return Worksheet.RangeRow(new XLRangeAddress(firstCellAddress, lastCellAddress));
        }

        /*public void Dispose()
        {
            // Dispose does nothing but left for not breaking the existing code
        }*/

        public IXLDataValidation SetDataValidation()
        {
            var existingValidation = GetDataValidation();
            if (existingValidation != null) return existingValidation;

            IXLDataValidation dataValidationToCopy = null;
            var dvEmpty = new List<IXLDataValidation>();
            foreach (IXLDataValidation dv in Worksheet.DataValidations)
            {
                foreach (IXLRange dvRange in dv.Ranges.GetIntersectedRanges(RangeAddress).ToList())
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
                                dv.Ranges.Add(Worksheet.Column(column.ColumnNumber()).Column(dvStart, thisStart - 1));
                                dv.Ranges.Add(Worksheet.Column(column.ColumnNumber()).Column(thisEnd + 1, dvEnd));
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
                                        dv.Ranges.Add(Worksheet.Column(column.ColumnNumber()).Column(coStart, coEnd));
                                    }
                                }
                            }
                        }
                        else
                        {
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

        public IXLConditionalFormat AddConditionalFormat()
        {
            var cf = new XLConditionalFormat(AsRange());
            Worksheet.ConditionalFormats.Add(cf);
            return cf;
        }

        internal IXLConditionalFormat AddConditionalFormat(IXLConditionalFormat source)
        {
            var cf = new XLConditionalFormat(AsRange());
            cf.CopyFrom(source);
            Worksheet.ConditionalFormats.Add(cf);
            return cf;
        }

        public void Select()
        {
            Worksheet.SelectedRanges.Add(AsRange());
        }

        public IXLRangeBase Grow()
        {
            return Grow(1);
        }

        public IXLRangeBase Grow(int growCount)
        {
            var firstRow = Math.Max(1, this.RangeAddress.FirstAddress.RowNumber - growCount);
            var firstColumn = Math.Max(1, this.RangeAddress.FirstAddress.ColumnNumber - growCount);

            var lastRow = Math.Min(XLHelper.MaxRowNumber, this.RangeAddress.LastAddress.RowNumber + growCount);
            var lastColumn = Math.Min(XLHelper.MaxColumnNumber, this.RangeAddress.LastAddress.ColumnNumber + growCount);

            return this.Worksheet.Range(firstRow, firstColumn, lastRow, lastColumn);
        }

        public IXLRangeBase Shrink()
        {
            return Shrink(1);
        }

        public IXLRangeBase Shrink(int shrinkCount)
        {
            var firstRow = this.RangeAddress.FirstAddress.RowNumber + shrinkCount;
            var firstColumn = this.RangeAddress.FirstAddress.ColumnNumber + shrinkCount;

            var lastRow = this.RangeAddress.LastAddress.RowNumber - shrinkCount;
            var lastColumn = this.RangeAddress.LastAddress.ColumnNumber - shrinkCount;

            if (firstRow > lastRow || firstColumn > lastColumn)
                return null;

            return this.Worksheet.Range(firstRow, firstColumn, lastRow, lastColumn);
        }

        public IXLRangeBase Intersection(IXLRangeBase otherRange, Func<IXLCell, Boolean> thisRangePredicate = null, Func<IXLCell, Boolean> otherRangePredicate = null)
        {
            if (otherRange == null)
                return null;

            if (!this.Worksheet.Equals(otherRange.Worksheet))
                return null;

            if (thisRangePredicate == null) thisRangePredicate = c => true;
            if (otherRangePredicate == null) otherRangePredicate = c => true;

            var intersectionCells = this.Cells(c => thisRangePredicate(c) && otherRange.Cells(otherRangePredicate).Contains(c));

            if (!intersectionCells.Any())
                return null;

            var firstRow = intersectionCells.Min(c => c.Address.RowNumber);
            var firstColumn = intersectionCells.Min(c => c.Address.ColumnNumber);

            var lastRow = intersectionCells.Max(c => c.Address.RowNumber);
            var lastColumn = intersectionCells.Max(c => c.Address.ColumnNumber);

            return this.Worksheet.Range(firstRow, firstColumn, lastRow, lastColumn);
        }

        public IXLCells SurroundingCells(Func<IXLCell, Boolean> predicate = null)
        {
            var cells = new XLCells(false, XLCellsUsedOptions.AllContents, predicate);
            this.Grow().Cells(c => !this.Contains(c)).ForEach(c => cells.Add(c as XLCell));
            return cells;
        }

        public IXLCells Union(IXLRangeBase otherRange, Func<IXLCell, Boolean> thisRangePredicate = null, Func<IXLCell, Boolean> otherRangePredicate = null)
        {
            if (otherRange == null)
                return this.Cells(thisRangePredicate);

            var cells = new XLCells(false, XLCellsUsedOptions.AllContents);
            if (!this.Worksheet.Equals(otherRange.Worksheet))
                return cells;

            if (thisRangePredicate == null) thisRangePredicate = c => true;
            if (otherRangePredicate == null) otherRangePredicate = c => true;

            this.Cells(thisRangePredicate).Concat(otherRange.Cells(otherRangePredicate)).Distinct().ForEach(c => cells.Add(c as XLCell));
            return cells;
        }

        public IXLCells Difference(IXLRangeBase otherRange, Func<IXLCell, Boolean> thisRangePredicate = null, Func<IXLCell, Boolean> otherRangePredicate = null)
        {
            if (otherRange == null)
                return this.Cells(thisRangePredicate);

            var cells = new XLCells(false, XLCellsUsedOptions.AllContents);
            if (!this.Worksheet.Equals(otherRange.Worksheet))
                return cells;

            if (thisRangePredicate == null) thisRangePredicate = c => true;
            if (otherRangePredicate == null) otherRangePredicate = c => true;

            this.Cells(c => thisRangePredicate(c) && !otherRange.Cells(otherRangePredicate).Contains(c)).ForEach(c => cells.Add(c as XLCell));
            return cells;
        }
    }
}
