#nullable disable

using System;
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
            Worksheet.RelocateRange(RangeType, oldAddress, newAddress);
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

        internal XLSheetRange SheetRange
        {
            get
            {
                if (!RangeAddress.IsValid)
                    throw new InvalidOperationException("Range address is invalid.");

                return XLSheetRange.FromRangeAddress(RangeAddress);
            }
        }

        public IXLDataValidation CreateDataValidation()
        {
            var newRange = AsRange();
            var dataValidation = new XLDataValidation(newRange);
            Worksheet.DataValidations.Add(dataValidation);
            return dataValidation;
        }

        public IXLDataValidation GetDataValidation()
        {
            Worksheet.DataValidations.TryGet(RangeAddress, out var existingDataValidation);
            return existingDataValidation;
        }

        #region IXLRangeBase Members

        IXLRangeAddress IXLAddressable.RangeAddress
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

        public String FormulaArrayA1
        {
            set
            {
                value = value?.TrimFormulaEqual();
                var range = XLSheetRange.FromRangeAddress(RangeAddress);
                if (Worksheet.MergedRanges.Any(mr => mr.Intersects(this)))
                    throw new InvalidOperationException("Can't create array function over a merged range.");

                if (Worksheet.Tables.Any<XLTable>(t => t.Intersects(this)))
                    throw new InvalidOperationException("Can't create array function over a table.");

                if (Cells(false).Any<XLCell>(c => c.HasArrayFormula && !RangeAddress.ContainsWhole(c.FormulaReference)))
                    throw new InvalidOperationException("Can't create array function that partially covers another array function.");

                var arrayFormula = XLCellFormula.Array(value, range, false);

                var formulaSlice = Worksheet.Internals.CellsCollection.FormulaSlice;
                formulaSlice.SetArray(range, arrayFormula);

                // If formula evaluates to a text, it is stored directly in a worksheet, not in SST. Thus
                // when the switch to formula happens, disable shared string and enable when formula is removed.
                var valueSlice = Worksheet.Internals.CellsCollection.ValueSlice;
                for (var row = range.TopRow; row <= range.BottomRow; ++row)
                {
                    for (var col = range.LeftColumn; col <= range.RightColumn; ++col)
                    {
                        valueSlice.SetShareString(new XLSheetPoint(row, col), false);
                    }
                }

                // Formula is shared across all cells, so it's enough to invalidate master cell
                var masterCell = FirstCell();
                masterCell.InvalidateFormula();
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

        public XLCellValue Value
        {
            set { Cells().ForEach(c => c.Value = value); }
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

        #endregion IXLStylized Members

        #endregion Public properties

        #region IXLRangeBase Members

        IXLCells IXLRangeBase.Cells(String cells) => Cells(cells);

        IXLCells IXLRangeBase.Cells(Boolean usedCellsOnly) => Cells(usedCellsOnly);

        IXLCells IXLRangeBase.Cells(Boolean usedCellsOnly, XLCellsUsedOptions options) => Cells(usedCellsOnly, options);

        IXLCells IXLRangeBase.CellsUsed() => CellsUsed();

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
            return FirstCellUsed(XLCellsUsedOptions.AllContents);
        }
        IXLCell IXLRangeBase.FirstCellUsed(XLCellsUsedOptions options)
        {
            return FirstCellUsed(options, null);
        }

        IXLCell IXLRangeBase.FirstCellUsed(Func<IXLCell, Boolean> predicate)
        {
            return FirstCellUsed(predicate);
        }

        IXLCell IXLRangeBase.FirstCellUsed(XLCellsUsedOptions options, Func<IXLCell, Boolean> predicate)
        {
            return FirstCellUsed(options, predicate);
        }

        IXLCell IXLRangeBase.LastCellUsed()
        {
            return LastCellUsed(XLCellsUsedOptions.AllContents);
        }
        IXLCell IXLRangeBase.LastCellUsed(XLCellsUsedOptions options)
        {
            return LastCellUsed(options, null);
        }

        IXLCell IXLRangeBase.LastCellUsed(Func<IXLCell, Boolean> predicate)
        {
            return LastCellUsed(predicate);
        }

        IXLCell IXLRangeBase.LastCellUsed(XLCellsUsedOptions options, Func<IXLCell, Boolean> predicate)
        {
            return LastCellUsed(options, predicate);
        }

        public virtual IXLCells Cells()
        {
            return Cells(false);
        }

        public virtual XLCells Cells(Boolean usedCellsOnly)
        {
            return Cells(usedCellsOnly, XLCellsUsedOptions.AllContents);
        }

        public XLCells Cells(Boolean usedCellsOnly, XLCellsUsedOptions options)
        {
            var cells = new XLCells(usedCellsOnly, options) { RangeAddress };
            return cells;
        }

        public virtual XLCells Cells(String cells)
        {
            return Ranges(cells).Cells();
        }

        public IXLCells Cells(Func<IXLCell, Boolean> predicate)
        {
            var cells = new XLCells(false, XLCellsUsedOptions.AllContents, predicate) { RangeAddress };
            return cells;
        }

        public XLCells CellsUsed()
        {
            return Cells(true);
        }

        public IXLRange Merge()
        {
            return Merge(true);
        }

        public IXLRange Merge(Boolean checkIntersect)
        {
            if (RangeAddress.FirstAddress.Equals(RangeAddress.LastAddress))
                return Worksheet.Range(RangeAddress);

            var asRange = AsRange();

            if (checkIntersect)
            {
                var intersectedMergedRanges =
                    Worksheet.Internals.MergedRanges.GetIntersectedRanges(RangeAddress).ToList();
                foreach (var intersectedMergedRange in intersectedMergedRanges)
                {
                    Worksheet.Internals.MergedRanges.Remove(intersectedMergedRange);
                }

                var firstCell = FirstCell();
                var firstCellStyleKey = (firstCell.Style as XLStyle).Key;
                var firstCellStyle = firstCell.Style;
                var defaultStyleKey = XLStyle.Default.Key;
                var cellsUsed =
                    CellsUsed(XLCellsUsedOptions.All & ~XLCellsUsedOptions.MergedRanges, c => !c.Equals(firstCell)).ToList();
                cellsUsed.ForEach(c => c.Clear(XLClearOptions.All
                                               & ~XLClearOptions.MergedRanges
                                               & ~XLClearOptions.NormalFormats));

                if (firstCellStyleKey.Alignment != defaultStyleKey.Alignment)
                    asRange.Style.Alignment = firstCellStyle.Alignment;
                else
                    cellsUsed.ForEach(c => c.Style.Alignment = firstCellStyle.Alignment);

                if (firstCellStyleKey.Fill != defaultStyleKey.Fill)
                    asRange.Style.Fill = firstCellStyle.Fill;
                else
                    cellsUsed.ForEach(c => c.Style.Fill = firstCellStyle.Fill);

                if (firstCellStyleKey.Font != defaultStyleKey.Font)
                    asRange.Style.Font = firstCellStyle.Font;
                else
                    cellsUsed.ForEach(c => c.Style.Font = firstCellStyle.Font);

                if (firstCellStyleKey.IncludeQuotePrefix != defaultStyleKey.IncludeQuotePrefix)
                    asRange.Style.IncludeQuotePrefix = firstCellStyle.IncludeQuotePrefix;
                else
                    cellsUsed.ForEach(c => c.Style.IncludeQuotePrefix = firstCellStyle.IncludeQuotePrefix);

                if (firstCellStyleKey.NumberFormat != defaultStyleKey.NumberFormat)
                    asRange.Style.NumberFormat = firstCellStyle.NumberFormat;
                else
                    cellsUsed.ForEach(c => c.Style.NumberFormat = firstCellStyle.NumberFormat);

                if (firstCellStyleKey.Protection != defaultStyleKey.Protection)
                    asRange.Style.Protection = firstCellStyle.Protection;
                else
                    cellsUsed.ForEach(c => c.Style.Protection = firstCellStyle.Protection);

                if (cellsUsed.Any(c => (c.Style as XLStyle).Key.Border != defaultStyleKey.Border))
                    asRange.Style.Border.SetInsideBorder(XLBorderStyleValues.None);
            }

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
            var cellClearOptions = clearOptions
                    & ~XLClearOptions.ConditionalFormats
                    & ~XLClearOptions.DataValidation
                    & ~XLClearOptions.MergedRanges
                    & ~XLClearOptions.Sparklines;
            var cellUsedOptions = cellClearOptions.ToCellsUsedOptions();
            foreach (var cell in CellsUsed(cellUsedOptions))
            {
                // We'll clear the conditional formatting, data validations, sparklines
                // and merged ranges later down.
                (cell as XLCell).Clear(cellClearOptions, true);
            }

            if (clearOptions.HasFlag(XLClearOptions.ConditionalFormats))
                RemoveConditionalFormatting();

            if (clearOptions.HasFlag(XLClearOptions.DataValidation))
            {
                var validation = CreateDataValidation();
                Worksheet.DataValidations.Delete(validation);
            }

            if (clearOptions.HasFlag(XLClearOptions.MergedRanges))
                ClearMerged();

            if (clearOptions.HasFlag(XLClearOptions.Sparklines))
                RemoveSparklines();

            if (clearOptions == XLClearOptions.All)
            {
                Worksheet.Internals.CellsCollection.Clear(XLSheetRange.FromRangeAddress(RangeAddress));
            }
            return this;
        }

        public IXLRangeBase Relative(IXLRangeBase sourceBaseRange, IXLRangeBase targetBaseRange)
        {
            var xlSourceBaseRangeAddress = (XLRangeAddress)sourceBaseRange.RangeAddress;
            var xlTargetBaseRangeAddress = (XLRangeAddress)targetBaseRange.RangeAddress;
            var xlRangeAddress = this.RangeAddress.Relative(in xlSourceBaseRangeAddress, in xlTargetBaseRangeAddress);

            return ((XLRangeBase)targetBaseRange).Range(in xlRangeAddress);
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

        internal void RemoveSparklines()
        {
            Worksheet.SparklineGroups.GetSparklines(this).ToList()
                .ForEach(sl => Worksheet.SparklineGroups.Remove(sl.Location));
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
            return Worksheet.Range(RangeAddress);
        }

        public IXLRange AddToNamed(String name)
        {
            return AddToNamed(name, XLScope.Workbook);
        }

        public IXLRange AddToNamed(String name, XLScope scope)
        {
            return AddToNamed(name, scope, null);
        }

        public IXLRange AddToNamed(String name, XLScope scope, String comment)
        {
            var definedNames = scope == XLScope.Workbook
                                  ? Worksheet.Workbook.DefinedNamesInternal
                                  : Worksheet.DefinedNames;

            if (definedNames.TryGetScopedValue(name, out var definedName))
                definedName.Add(RangeAddress.ToStringFixed(XLReferenceStyle.A1, true));
            else
                definedNames.Add(name, RangeAddress.ToStringFixed(XLReferenceStyle.A1, true), comment);

            return AsRange();
        }

        public IXLRangeBase SetValue(XLCellValue value)
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
            return !CellsUsed().Any<XLCell>() || CellsUsed().Any<XLCell>(c => c.IsEmpty());
        }

        public virtual Boolean IsEmpty(XLCellsUsedOptions options)
        {
            foreach (var cell in CellsUsed(options))
            {
                if (!cell.IsEmpty(options))
                {
                    return false;
                }
            }
            return true;
        }

        public virtual Boolean IsEntireRow()
        {
            return RangeAddress.IsEntireRow();
        }

        public virtual Boolean IsEntireColumn()
        {
            return RangeAddress.IsEntireColumn();
        }

        public Boolean IsEntireSheet()
        {
            return RangeAddress.IsEntireSheet();
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
                               || culture.CompareInfo.IndexOf(c.Value.ToString(CultureInfo.CurrentCulture), searchText, compareOptions) >= 0;
                    else
                        return culture.CompareInfo.IndexOf(c.Value.ToString(CultureInfo.CurrentCulture), searchText, compareOptions) >= 0;
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

        internal XLCell FirstCellUsed(Func<IXLCell, Boolean> predicate)
        {
            return FirstCellUsed(XLCellsUsedOptions.AllContents, predicate);
        }

        internal XLCell FirstCellUsed(XLCellsUsedOptions options, Func<IXLCell, Boolean> predicate = null)
        {
            var cellsUsed = CellsUsedInternal(options, r => r.FirstCell(), predicate).ToList();

            if (!cellsUsed.Any())
                return null;

            var firstRow = cellsUsed.Min(c => c.Address.RowNumber);
            var firstColumn = cellsUsed.Min(c => c.Address.ColumnNumber);

            if (firstRow < RangeAddress.FirstAddress.RowNumber)
                firstRow = RangeAddress.FirstAddress.RowNumber;

            if (firstColumn < RangeAddress.FirstAddress.ColumnNumber)
                firstColumn = RangeAddress.FirstAddress.ColumnNumber;

            return Worksheet.Cell(firstRow, firstColumn);
        }

        internal XLCell LastCellUsed()
        {
            return LastCellUsed(XLCellsUsedOptions.AllContents, predicate: null);
        }

        internal XLCell LastCellUsed(Func<IXLCell, Boolean> predicate)
        {
            return LastCellUsed(XLCellsUsedOptions.AllContents, predicate);
        }

        internal XLCell LastCellUsed(XLCellsUsedOptions options, Func<IXLCell, Boolean> predicate = null)
        {
            var cellsUsed = CellsUsedInternal(options, r => r.LastCell(), predicate).ToList();

            if (!cellsUsed.Any())
                return null;

            var lastRow = cellsUsed.Max(c => c.Address.RowNumber);
            var lastColumn = cellsUsed.Max(c => c.Address.ColumnNumber);

            if (lastRow > RangeAddress.LastAddress.RowNumber)
                lastRow = RangeAddress.LastAddress.RowNumber;

            if (lastColumn > RangeAddress.LastAddress.ColumnNumber)
                lastColumn = RangeAddress.LastAddress.ColumnNumber;

            return Worksheet.Cell(lastRow, lastColumn);
        }

        public XLCell Cell(Int32 row, Int32 column)
        {
            return Cell(new XLAddress(Worksheet, row, column, false, false));
        }

        public virtual XLCell Cell(String cellAddressInRange)
        {
            if (XLHelper.IsValidA1Address(cellAddressInRange))
                return Cell(XLAddress.Create(Worksheet, cellAddressInRange));

            if (Worksheet.DefinedNames.TryGetValue(cellAddressInRange, out IXLDefinedName definedName))
                return definedName.Ranges.First().FirstCell().CastTo<XLCell>();

            return null;
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

            var cell = Worksheet.Internals.CellsCollection.GetCell(new XLSheetPoint(absRow, absColumn));
            return cell;
        }

        public Int32 RowCount()
        {
            return RangeAddress.LastAddress.RowNumber - RangeAddress.FirstAddress.RowNumber + 1;
        }

        public Int32 RowCount(XLCellsUsedOptions cellsUsedOptions)
        {
            var lcu = LastCellUsed(cellsUsedOptions);
            if (lcu == null) return 0;

            var fcu = FirstCellUsed(cellsUsedOptions);
            if (fcu == null) return 0;

            return lcu.Address.RowNumber - fcu.Address.RowNumber + 1;
        }

        public Int32 RowNumber()
        {
            return RangeAddress.FirstAddress.RowNumber;
        }

        public Int32 ColumnCount()
        {
            return RangeAddress.LastAddress.ColumnNumber - RangeAddress.FirstAddress.ColumnNumber + 1;
        }

        public Int32 ColumnCount(XLCellsUsedOptions cellsUsedOptions)
        {
            var lcu = LastCellUsed(cellsUsedOptions);
            if (lcu == null) return 0;

            var fcu = FirstCellUsed(cellsUsedOptions);
            if (fcu == null) return 0;

            return lcu.Address.ColumnNumber - fcu.Address.ColumnNumber + 1;
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
            if (!this.Worksheet.Equals(newFirstCellAddress.Worksheet))
                throw new ArgumentException("The address refers to a different worksheet.", nameof(newFirstCellAddress));

            if (!this.Worksheet.Equals(newLastCellAddress.Worksheet))
                throw new ArgumentException("The address refers to a different worksheet.", nameof(newLastCellAddress));

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
            var rangeAddress = new XLRangeAddress
            (
                new XLAddress
                (
                    Worksheet,
                    firstCellRow + RangeAddress.FirstAddress.RowNumber - 1,
                    firstCellColumn + RangeAddress.FirstAddress.ColumnNumber - 1,
                    fixedRow: false,
                    fixedColumn: false
                ),
                new XLAddress
                (
                    Worksheet,
                    lastCellRow + RangeAddress.FirstAddress.RowNumber - 1,
                    lastCellColumn + RangeAddress.FirstAddress.ColumnNumber - 1,
                    fixedRow: false,
                    fixedColumn: false
                )
            );
            return Range(rangeAddress);
        }

        public XLRange Range(IXLAddress firstCellAddress, IXLAddress lastCellAddress)
        {
            var rangeAddress = new XLRangeAddress((XLAddress)firstCellAddress, (XLAddress)lastCellAddress);
            return Range(rangeAddress);
        }

        public XLRange Range(IXLRangeAddress rangeAddress)
        {
            var xlRangeAddress = (XLRangeAddress)rangeAddress;
            return Range(in xlRangeAddress);
        }

        internal XLRange Range(in XLRangeAddress rangeAddress)
        {
            var ws = rangeAddress.FirstAddress.Worksheet ??
                     rangeAddress.LastAddress.Worksheet ??
                     Worksheet;

            var newFirstCellAddress = new XLAddress(ws,
                                 rangeAddress.FirstAddress.RowNumber,
                                 rangeAddress.FirstAddress.ColumnNumber,
                                 rangeAddress.FirstAddress.FixedRow,
                                 rangeAddress.FirstAddress.FixedColumn);

            var newLastCellAddress = new XLAddress(ws,
                                rangeAddress.LastAddress.RowNumber,
                                rangeAddress.LastAddress.ColumnNumber,
                                rangeAddress.LastAddress.FixedRow,
                                rangeAddress.LastAddress.FixedColumn);

            return GetRange(newFirstCellAddress, newLastCellAddress);
        }

        public virtual XLRanges Ranges(String ranges)
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
            if (Int32.TryParse(address, out Int32 rowNumber))
                return RangeAddress.FirstAddress.ColumnLetter + (rowNumber + RangeAddress.FirstAddress.RowNumber - 1).ToInvariantString();
            return address;
        }

        protected String FixRowAddress(String address)
        {
            if (Int32.TryParse(address, out Int32 columnNumber))
                return XLHelper.GetColumnLetterFromNumber(columnNumber + RangeAddress.FirstAddress.ColumnNumber - 1) + RangeAddress.FirstAddress.RowNumber.ToInvariantString();
            return address;
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
                foreach (XLCell cell in ws.Internals.CellsCollection.GetCells(c => c.Formula is not null))
                    cell.ShiftFormulaColumns(AsRange(), numberOfColumns);
            }

            Worksheet.SparklineGroupsInternal.ShiftColumns(XLSheetRange.FromRangeAddress(RangeAddress), numberOfColumns);

            // Inserting and shifting of whole columns is rather inconsistent across the codebase. In some places, the columns collection
            // is shifted before this method is called and thus the we can't shift column properties again. In others, the code relies on
            // shifting in this method.
            if (!onlyUsedCells)
            {
                int lastColumn = Worksheet.Internals.CellsCollection.MaxColumnUsed;
                if (lastColumn > 0)
                {
                    int firstColumn = RangeAddress.FirstAddress.ColumnNumber;
                    for (int co = lastColumn; co >= firstColumn; co--)
                    {
                        int newColumn = co + numberOfColumns;
                        if (this.IsEntireColumn())
                        {
                            Worksheet.Column(newColumn).Width = Worksheet.Column(co).Width;
                        }
                    }
                }
            }

            var insertedRange = new XLSheetRange(
                XLSheetPoint.FromAddress(RangeAddress.FirstAddress),
                new XLSheetPoint(RangeAddress.LastAddress.RowNumber, RangeAddress.FirstAddress.ColumnNumber + numberOfColumns - 1));

            Worksheet.Internals.CellsCollection.InsertAreaAndShiftRight(insertedRange);

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
                        var styleToUse =
                            Worksheet.Internals.RowsCollection.TryGetValue(ro, out XLRow row)
                                ? row.Style
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
                foreach (XLCell cell in ws.Internals.CellsCollection.GetCells(c => c.Formula is not null))
                    cell.ShiftFormulaRows(asRange, numberOfRows);
            }

            Worksheet.SparklineGroupsInternal.ShiftRows(XLSheetRange.FromRangeAddress(RangeAddress), numberOfRows);

            if (!onlyUsedCells)
            {
                int lastRow = Worksheet.Internals.CellsCollection.MaxRowUsed;
                if (lastRow > 0)
                {
                    int firstRow = RangeAddress.FirstAddress.RowNumber;
                    for (var ro = lastRow; ro >= firstRow; ro--)
                    {
                        var newRow = ro + numberOfRows;
                        if (this.IsEntireRow())
                        {
                            Worksheet.Row(newRow).Height = Worksheet.Row(ro).Height;
                        }
                    }
                }
            }

            var insertedRange = new XLSheetRange(
                XLSheetPoint.FromAddress(RangeAddress.FirstAddress),
                new XLSheetPoint(RangeAddress.FirstAddress.RowNumber + numberOfRows - 1, RangeAddress.LastAddress.ColumnNumber));
            Worksheet.Internals.CellsCollection.InsertAreaAndShiftDown(insertedRange);

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
                        var styleToUse =
                            Worksheet.Internals.ColumnsCollection.TryGetValue(co, out XLColumn column)
                                ? column.Style
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

            Worksheet.SparklineGroups.Remove(this);

            IXLRange shiftedRangeFormula = Worksheet.Range(
                RangeAddress.FirstAddress.RowNumber,
                RangeAddress.FirstAddress.ColumnNumber,
                RangeAddress.LastAddress.RowNumber,
                RangeAddress.LastAddress.ColumnNumber);

            // Shift formulas first
            foreach (var cell in Worksheet
                .Workbook
                .Worksheets
                .Cast<XLWorksheet>()
                .SelectMany(ws => ws
                    .Internals
                    .CellsCollection
                    .GetCells(c => c.HasFormula)))
            {
                if (shiftDeleteCells == XLShiftDeletedCells.ShiftCellsUp)
                    cell.ShiftFormulaRows((XLRange)shiftedRangeFormula, numberOfRows * -1);
                else
                    cell.ShiftFormulaColumns((XLRange)shiftedRangeFormula, numberOfColumns * -1);
            }

            // Range to shift...
            Int32 columnModifier = 0;
            Int32 rowModifier = 0;
            var range = XLSheetRange.FromRangeAddress(RangeAddress);
            switch (shiftDeleteCells)
            {
                case XLShiftDeletedCells.ShiftCellsLeft:
                    Worksheet.Internals.CellsCollection.DeleteAreaAndShiftLeft(range);
                    Worksheet.SparklineGroupsInternal.ShiftColumns(range, -numberOfColumns);
                    columnModifier = ColumnCount();
                    break;

                case XLShiftDeletedCells.ShiftCellsUp:
                    Worksheet.Internals.CellsCollection.DeleteAreaAndShiftUp(range);
                    Worksheet.SparklineGroupsInternal.ShiftRows(range, -numberOfRows);
                    rowModifier = RowCount();
                    break;
            }

            var mergesToRemove = Worksheet.Internals.MergedRanges.Where(Contains).ToList();
            mergesToRemove.ForEach(r => Worksheet.Internals.MergedRanges.Remove(r));

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
            var lastAddress = (XLAddress)thisRangeAddress.LastAddress;

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
            CopyTo((XLCell)target.FirstCell());
        }

        internal void CopyTo(XLCell target)
        {
            target.CopyFrom(this);
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

        public virtual IXLAutoFilter SetAutoFilter()
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

            SortColumns.CastTo<XLSortElements>().AddRange(ParseSortOrder(columnsToSortBy, sortOrder, matchCase, ignoreBlanks));

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

        private void SortRangeRows()
        {
            var sortRange = SheetRange;
            var cellsCollection = Worksheet.Internals.CellsCollection;
            if (sortRange.IsEntireColumn())
            {
                // If we're dealing with the entire column, we're not interested in the unused cells
                var lastRowUsed = cellsCollection.LastRowUsed(XLSheetRange.Full, XLCellsUsedOptions.Contents);
                sortRange = new XLSheetRange(sortRange.FirstPoint, new XLSheetPoint(lastRowUsed, sortRange.RightColumn));
            }

            var comparer = new XLRangeRowsSortComparer(Worksheet, sortRange, SortColumns);
            var rows = new int[sortRange.Height];
            for (var i = 0; i < sortRange.Height; ++i)
                rows[i] = i + sortRange.TopRow;

            Array.Sort(rows, comparer);

            cellsCollection.RemapRows(rows, sortRange);
        }

        private void SortRangeColumns()
        {
            var sortRange = SheetRange;
            var cellsCollection = Worksheet.Internals.CellsCollection;
            if (sortRange.IsEntireRow())
            {
                // If we're dealing with the entire row, we're not interested in the unused cells
                var lastColumnCell = cellsCollection.LastColumnUsed(XLSheetRange.Full, XLCellsUsedOptions.Contents);
                sortRange = new XLSheetRange(sortRange.FirstPoint, new XLSheetPoint(sortRange.BottomRow, lastColumnCell));
            }

            var comparer = new XLRangeColumnsSortComparer(Worksheet, sortRange, SortRows);
            var columns = new int[sortRange.Width];
            for (var i = 0; i < sortRange.Width; ++i)
                columns[i] = i + sortRange.LeftColumn;

            Array.Sort(columns, comparer);

            cellsCollection.RemapColumns(columns, sortRange);
        }

        private IEnumerable<XLSortElement> ParseSortOrder(string columnsToSortBy, XLSortOrder defaultSortOrder, bool matchCase, bool ignoreBlanks)
        {
            foreach (var sortColumn in columnsToSortBy.Split(',').Select(coPair => coPair.Trim()))
            {
                var sortOrder = defaultSortOrder;

                String columnName;
                if (sortColumn.Contains(' '))
                {
                    var pair = sortColumn.Split(' ');
                    columnName = pair[0];
                    sortOrder = pair[1].Equals("ASC", StringComparison.OrdinalIgnoreCase) ? XLSortOrder.Ascending : XLSortOrder.Descending;
                }
                else
                {
                    columnName = sortColumn;
                }

                if (!Int32.TryParse(columnName, out Int32 columnNumber))
                    columnNumber = XLHelper.GetColumnNumberFromLetter(columnName);

                yield return new XLSortElement(
                    columnNumber,
                    sortOrder,
                    ignoreBlanks,
                    matchCase);
            }
        }

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

        [Obsolete("Use GetDataValidation() to access the existing rule, or CreateDataValidation() to create a new one.")]
        public IXLDataValidation SetDataValidation()
        {
            var existingValidation = GetDataValidation();
            if (existingValidation != null && existingValidation.Ranges.Any(r => r == this))
                return existingValidation;

            IXLDataValidation dataValidationToCopy = Worksheet.DataValidations.GetAllInRange(RangeAddress)
                .FirstOrDefault();

            var newRange = AsRange();
            var dataValidation = new XLDataValidation(newRange);
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

        public IXLRangeAddress Intersection(IXLRangeBase otherRange, Func<IXLCell, Boolean> thisRangePredicate = null, Func<IXLCell, Boolean> otherRangePredicate = null)
        {
            if (otherRange == null)
                return null;

            if (!this.Worksheet.Equals(otherRange.Worksheet))
                return null;

            if (thisRangePredicate == null && otherRangePredicate == null)
            {
                // Special case, no predicates. We can optimise this a bit then.
                return this.RangeAddress.Intersection(otherRange.RangeAddress);
            }
            else
            {
                thisRangePredicate = thisRangePredicate ?? (c => true);
                otherRangePredicate = otherRangePredicate ?? (c => true);

                var intersectionCells = this.Cells(c => thisRangePredicate(c) && otherRange.Cells(otherRangePredicate).Contains(c));

                if (!intersectionCells.Any())
                    return null;

                var firstRow = intersectionCells.Min(c => c.Address.RowNumber);
                var firstColumn = intersectionCells.Min(c => c.Address.ColumnNumber);

                var lastRow = intersectionCells.Max(c => c.Address.RowNumber);
                var lastColumn = intersectionCells.Max(c => c.Address.ColumnNumber);

                return new XLRangeAddress
                (
                    new XLAddress(this.Worksheet, firstRow, firstColumn, fixedRow: false, fixedColumn: false),
                    new XLAddress(this.Worksheet, lastRow, lastColumn, fixedRow: false, fixedColumn: false)
                );
            }
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

        private IEnumerable<IXLCell> CellsUsedInternal(XLCellsUsedOptions options, Func<IXLRange, IXLCell> selector, Func<IXLCell, bool> predicate)
        {
            predicate ??= (t => true);

            //To avoid unnecessary initialization of thousands cells
            var opt = options
                      & ~XLCellsUsedOptions.ConditionalFormats
                      & ~XLCellsUsedOptions.DataValidation
                      & ~XLCellsUsedOptions.MergedRanges;

            // If opt == 0 then we're basically back at unconstrained, so just set back the original options
            if (opt == XLCellsUsedOptions.NoConstraints)
                opt = options;

            IEnumerable<IXLCell> cellsUsed = CellsUsed(opt, predicate);

            if (options.HasFlag(XLCellsUsedOptions.ConditionalFormats))
            {
                cellsUsed = cellsUsed.Union(
                    Worksheet.ConditionalFormats
                        .SelectMany(cf => cf.Ranges.GetIntersectedRanges(RangeAddress))
                        .Select(selector)
                        .Where(predicate)
                );
            }
            if (options.HasFlag(XLCellsUsedOptions.DataValidation))
            {
                cellsUsed = cellsUsed.Union(
                    Worksheet.DataValidations
                        .GetAllInRange(RangeAddress)
                        .SelectMany(dv => dv.Ranges)
                        .Select(selector)
                        .Where(predicate)
                );
            }
            if (options.HasFlag(XLCellsUsedOptions.MergedRanges))
            {
                cellsUsed = cellsUsed.Union(
                    Worksheet.MergedRanges.GetIntersectedRanges(RangeAddress)
                        .Select(selector)
                        .Where(predicate)
                );
            }

            return cellsUsed;
        }
    }
}
