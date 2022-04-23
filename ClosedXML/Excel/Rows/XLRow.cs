using SkiaSharp;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLRow : XLRangeBase, IXLRow
    {
        #region Private fields

        private double _height;
        private int _outlineLevel;

        #endregion Private fields

        #region Constructor

        /// <summary>
        /// The direct contructor should only be used in <see cref="XLWorksheet.RangeFactory"/>.
        /// </summary>
        public XLRow(XLWorksheet worksheet, int row)
            : base(XLRangeAddress.EntireRow(worksheet, row), worksheet.StyleValue)
        {
            SetRowNumber(row);

            _height = worksheet.RowHeight;
        }

        #endregion Constructor

        public override XLRangeType RangeType
        {
            get { return XLRangeType.Row; }
        }

        public override IEnumerable<IXLStyle> Styles
        {
            get
            {
                yield return Style;

                var row = RowNumber();

                foreach (var cell in Worksheet.Internals.CellsCollection.GetCellsInRow(row))
                    yield return cell.Style;
            }
        }

        protected override IEnumerable<XLStylizedBase> Children
        {
            get
            {
                var row = RowNumber();

                foreach (var cell in Worksheet.Internals.CellsCollection.GetCellsInRow(row))
                    yield return cell;
            }
        }

        public bool Collapsed { get; set; }

        #region IXLRow Members

        public bool Loading { get; set; }

        public bool HeightChanged { get; private set; }

        public double Height
        {
            get { return _height; }
            set
            {
                if (!Loading)
                    HeightChanged = true;

                _height = value;
            }
        }

        public void ClearHeight()
        {
            Height = Worksheet.RowHeight;
            HeightChanged = false;
        }

        public void Delete()
        {
            var rowNumber = RowNumber();
            AsRange().Delete(XLShiftDeletedCells.ShiftCellsUp);
            Worksheet.DeleteRow(rowNumber);
        }

        public new IXLRows InsertRowsBelow(int numberOfRows)
        {
            var rowNum = RowNumber();
            Worksheet.Internals.RowsCollection.ShiftRowsDown(rowNum + 1, numberOfRows);
            var asRange = Worksheet.Row(rowNum).AsRange();
            asRange.InsertRowsBelowVoid(true, numberOfRows);

            var newRows = Worksheet.Rows(rowNum + 1, rowNum + numberOfRows);

            CopyRows(newRows);

            return newRows;
        }

        private void CopyRows(IXLRows newRows)
        {
            foreach (var newRow in newRows)
            {
                var internalRow = Worksheet.Internals.RowsCollection[newRow.RowNumber()];
                internalRow._height = Height;
                internalRow.InnerStyle = InnerStyle;
                internalRow.Collapsed = Collapsed;
                internalRow.IsHidden = IsHidden;
                internalRow._outlineLevel = OutlineLevel;
            }
        }

        public new IXLRows InsertRowsAbove(int numberOfRows)
        {
            var rowNum = RowNumber();
            if (rowNum > 1)
            {
                return Worksheet.Row(rowNum - 1).InsertRowsBelow(numberOfRows);
            }

            Worksheet.Internals.RowsCollection.ShiftRowsDown(rowNum, numberOfRows);
            var asRange = Worksheet.Row(rowNum).AsRange();
            asRange.InsertRowsAboveVoid(true, numberOfRows);

            return Worksheet.Rows(rowNum, rowNum + numberOfRows - 1);
        }

        public new IXLRow Clear(XLClearOptions clearOptions = XLClearOptions.All)
        {
            base.Clear(clearOptions);
            return this;
        }

        public IXLCell Cell(int columnNumber)
        {
            return Cell(1, columnNumber);
        }

        public override XLCell Cell(string columnLetter)
        {
            return Cell(1, columnLetter);
        }

        IXLCell IXLRow.Cell(string columnLetter)
        {
            return Cell(columnLetter);
        }

        public override IXLCells Cells()
        {
            return Cells(true, XLCellsUsedOptions.All);
        }

        public override IXLCells Cells(bool usedCellsOnly)
        {
            if (usedCellsOnly)
                return Cells(true, XLCellsUsedOptions.AllContents);
            else
                return Cells(FirstCellUsed().Address.ColumnNumber, LastCellUsed().Address.ColumnNumber);
        }

        public override IXLCells Cells(string cellsInRow)
        {
            var retVal = new XLCells(false, XLCellsUsedOptions.AllContents);
            var rangePairs = cellsInRow.Split(',');
            foreach (var pair in rangePairs)
                retVal.Add(Range(pair.Trim()).RangeAddress);
            return retVal;
        }

        public IXLCells Cells(int firstColumn, int lastColumn)
        {
            return Cells(firstColumn + ":" + lastColumn);
        }

        public IXLCells Cells(string firstColumn, string lastColumn)
        {
            return Cells(XLHelper.GetColumnNumberFromLetter(firstColumn) + ":"
                         + XLHelper.GetColumnNumberFromLetter(lastColumn));
        }

        public IXLRow AdjustToContents(int startColumn)
        {
            return AdjustToContents(startColumn, XLHelper.MaxColumnNumber);
        }

        public IXLRow AdjustToContents(int startColumn, int endColumn)
        {
            return AdjustToContents(startColumn, endColumn, 0, double.MaxValue);
        }

        public IXLRow AdjustToContents(double minHeight, double maxHeight)
        {
            return AdjustToContents(1, XLHelper.MaxColumnNumber, minHeight, maxHeight);
        }

        public IXLRow AdjustToContents(int startColumn, double minHeight, double maxHeight)
        {
            return AdjustToContents(startColumn, XLHelper.MaxColumnNumber, minHeight, maxHeight);
        }

        public IXLRow AdjustToContents(int startColumn, int endColumn, double minHeight, double maxHeight)
        {
            var fontCache = new Dictionary<IXLFontBase, SKFont>();

            var rowMaxHeight = minHeight;
            foreach (var c in from XLCell c in Row(startColumn, endColumn).CellsUsed() where !c.IsMerged() select c)
            {
                double thisHeight;
                var textRotation = c.StyleValue.Alignment.TextRotation;
                if (c.HasRichText || textRotation != 0 || c.InnerText.Contains(XLConstants.NewLine))
                {
                    var kpList = new List<KeyValuePair<IXLFontBase, string>>();
                    if (c.HasRichText)
                    {
                        foreach (var rt in c.GetRichText())
                        {
                            var formattedString = rt.Text;
                            var arr = formattedString.Split(new[] { XLConstants.NewLine }, StringSplitOptions.None);
                            var arrCount = arr.Count();
                            for (var i = 0; i < arrCount; i++)
                            {
                                var s = arr[i];
                                if (i < arrCount - 1)
                                    s += XLConstants.NewLine;
                                kpList.Add(new KeyValuePair<IXLFontBase, string>(rt, s));
                            }
                        }
                    }
                    else
                    {
                        var formattedString = c.GetFormattedString();
                        var arr = formattedString.Split(new[] { XLConstants.NewLine }, StringSplitOptions.None);
                        var arrCount = arr.Count();
                        for (var i = 0; i < arrCount; i++)
                        {
                            var s = arr[i];
                            if (i < arrCount - 1)
                                s += XLConstants.NewLine;
                            kpList.Add(new KeyValuePair<IXLFontBase, string>(c.Style.Font, s));
                        }
                    }

                    double maxLongCol = kpList.Max(kp => kp.Value.Length);
                    var maxHeightCol = kpList.Max(kp => kp.Key.GetHeight(fontCache));
                    var lineCount = kpList.Count(kp => kp.Value.Contains(XLConstants.NewLine)) + 1;
                    if (textRotation == 0)
                        thisHeight = maxHeightCol * lineCount;
                    else
                    {
                        if (textRotation == 255)
                            thisHeight = maxLongCol * maxHeightCol;
                        else
                        {
                            double rotation;
                            if (textRotation == 90 || textRotation == 180)
                                rotation = 90;
                            else
                                rotation = textRotation % 90;

                            thisHeight = rotation / 90.0 * maxHeightCol * maxLongCol * 0.5;
                        }
                    }
                }
                else
                    thisHeight = c.Style.Font.GetHeight(fontCache);

                if (thisHeight >= maxHeight)
                {
                    rowMaxHeight = maxHeight;
                    break;
                }
                if (thisHeight > rowMaxHeight)
                    rowMaxHeight = thisHeight;
            }

            if (rowMaxHeight <= 0)
                rowMaxHeight = Worksheet.RowHeight;

            Height = rowMaxHeight;

            foreach (IDisposable font in fontCache.Values)
            {
                font.Dispose();
            }
            return this;
        }

        public IXLRow Hide()
        {
            IsHidden = true;
            return this;
        }

        public IXLRow Unhide()
        {
            IsHidden = false;
            return this;
        }

        public bool IsHidden { get; set; }

        public int OutlineLevel
        {
            get { return _outlineLevel; }
            set
            {
                if (value < 0 || value > 8)
                    throw new ArgumentOutOfRangeException("value", "Outline level must be between 0 and 8.");

                Worksheet.IncrementColumnOutline(value);
                Worksheet.DecrementColumnOutline(_outlineLevel);
                _outlineLevel = value;
            }
        }

        public IXLRow Group()
        {
            return Group(false);
        }

        public IXLRow Group(int outlineLevel)
        {
            return Group(outlineLevel, false);
        }

        public IXLRow Ungroup()
        {
            return Ungroup(false);
        }

        public IXLRow Group(bool collapse)
        {
            if (OutlineLevel < 8)
                OutlineLevel += 1;

            Collapsed = collapse;
            return this;
        }

        public IXLRow Group(int outlineLevel, bool collapse)
        {
            OutlineLevel = outlineLevel;
            Collapsed = collapse;
            return this;
        }

        public IXLRow Ungroup(bool ungroupFromAll)
        {
            if (ungroupFromAll)
                OutlineLevel = 0;
            else
            {
                if (OutlineLevel > 0)
                    OutlineLevel -= 1;
            }
            return this;
        }

        public IXLRow Collapse()
        {
            Collapsed = true;
            return Hide();
        }

        public IXLRow Expand()
        {
            Collapsed = false;
            return Unhide();
        }

        public int CellCount()
        {
            return RangeAddress.LastAddress.ColumnNumber - RangeAddress.FirstAddress.ColumnNumber + 1;
        }

        public new IXLRow Sort()
        {
            return SortLeftToRight();
        }

        public new IXLRow SortLeftToRight(XLSortOrder sortOrder = XLSortOrder.Ascending, bool matchCase = false,
                                          bool ignoreBlanks = true)
        {
            base.SortLeftToRight(sortOrder, matchCase, ignoreBlanks);
            return this;
        }

        IXLRangeRow IXLRow.CopyTo(IXLCell target)
        {
            var copy = AsRange().CopyTo(target);
            return copy.Row(1);
        }

        IXLRangeRow IXLRow.CopyTo(IXLRangeBase target)
        {
            var copy = AsRange().CopyTo(target);
            return copy.Row(1);
        }

        public IXLRow CopyTo(IXLRow row)
        {
            row.Clear();
            var newRow = (XLRow)row;
            newRow._height = _height;
            newRow.HeightChanged = HeightChanged;
            newRow.InnerStyle = GetStyle();
            newRow.IsHidden = IsHidden;

            AsRange().CopyTo(row);

            return newRow;
        }

        public IXLRangeRow Row(int start, int end)
        {
            return Range(1, start, 1, end).Row(1);
        }

        public IXLRangeRow Row(IXLCell start, IXLCell end)
        {
            return Row(start.Address.ColumnNumber, end.Address.ColumnNumber);
        }

        public IXLRangeRows Rows(string rows)
        {
            var retVal = new XLRangeRows();
            var rowPairs = rows.Split(',');
            foreach (var pair in rowPairs)
                AsRange().Rows(pair.Trim()).ForEach(retVal.Add);

            return retVal;
        }

        public IXLRow AddHorizontalPageBreak()
        {
            Worksheet.PageSetup.AddHorizontalPageBreak(RowNumber());
            return this;
        }

        public IXLRow SetDataType(XLDataType dataType)
        {
            DataType = dataType;
            return this;
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        public IXLRangeRow RowUsed(bool includeFormats)
        {
            return RowUsed(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents);
        }

        public IXLRangeRow RowUsed(XLCellsUsedOptions options = XLCellsUsedOptions.AllContents)
        {
            return Row((this as IXLRangeBase).FirstCellUsed(options),
                (this as IXLRangeBase).LastCellUsed(options));
        }

        #endregion IXLRow Members

        public override XLRange AsRange()
        {
            return Range(1, 1, 1, XLHelper.MaxColumnNumber);
        }

        internal override void WorksheetRangeShiftedColumns(XLRange range, int columnsShifted)
        {
            //do nothing
        }

        internal override void WorksheetRangeShiftedRows(XLRange range, int rowsShifted)
        {
            return; // rows are shifted by XLRowCollection
        }

        internal void SetRowNumber(int row)
        {
            RangeAddress = new XLRangeAddress(
                new XLAddress(Worksheet, row, 1, RangeAddress.FirstAddress.FixedRow,
                              RangeAddress.FirstAddress.FixedColumn),
                new XLAddress(Worksheet,
                              row,
                              XLHelper.MaxColumnNumber,
                              RangeAddress.LastAddress.FixedRow,
                              RangeAddress.LastAddress.FixedColumn));
        }

        public override XLRange Range(string rangeAddressStr)
        {
            string rangeAddressToUse;
            if (rangeAddressStr.Contains(':') || rangeAddressStr.Contains('-'))
            {
                if (rangeAddressStr.Contains('-'))
                    rangeAddressStr = rangeAddressStr.Replace('-', ':');

                var arrRange = rangeAddressStr.Split(':');
                var firstPart = arrRange[0];
                var secondPart = arrRange[1];
                rangeAddressToUse = FixRowAddress(firstPart) + ":" + FixRowAddress(secondPart);
            }
            else
                rangeAddressToUse = FixRowAddress(rangeAddressStr);

            var rangeAddress = new XLRangeAddress(Worksheet, rangeAddressToUse);
            return Range(rangeAddress);
        }

        public IXLRow AdjustToContents()
        {
            return AdjustToContents(1);
        }

        internal void SetStyleNoColumns(IXLStyle value)
        {
            InnerStyle = value;

            var row = RowNumber();
            foreach (var c in Worksheet.Internals.CellsCollection.GetCellsInRow(row))
                c.InnerStyle = value;
        }

        private XLRow RowShift(int rowsToShift)
        {
            return Worksheet.Row(RowNumber() + rowsToShift);
        }

        #region XLRow Above

        IXLRow IXLRow.RowAbove()
        {
            return RowAbove();
        }

        IXLRow IXLRow.RowAbove(int step)
        {
            return RowAbove(step);
        }

        public XLRow RowAbove()
        {
            return RowAbove(1);
        }

        public XLRow RowAbove(int step)
        {
            return RowShift(step * -1);
        }

        #endregion XLRow Above

        #region XLRow Below

        IXLRow IXLRow.RowBelow()
        {
            return RowBelow();
        }

        IXLRow IXLRow.RowBelow(int step)
        {
            return RowBelow(step);
        }

        public XLRow RowBelow()
        {
            return RowBelow(1);
        }

        public XLRow RowBelow(int step)
        {
            return RowShift(step);
        }

        #endregion XLRow Below

        public override bool IsEmpty()
        {
            return IsEmpty(XLCellsUsedOptions.AllContents);
        }

        public override bool IsEmpty(XLCellsUsedOptions options)
        {
            if (options.HasFlag(XLCellsUsedOptions.NormalFormats) &&
                !StyleValue.Equals(Worksheet.StyleValue))
                return false;

            return base.IsEmpty(options);
        }

        public override bool IsEntireRow()
        {
            return true;
        }

        public override bool IsEntireColumn()
        {
            return false;
        }
    }
}