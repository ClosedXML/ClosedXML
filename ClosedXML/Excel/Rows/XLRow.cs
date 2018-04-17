using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLRow : XLRangeBase, IXLRow
    {
        #region Private fields

        private Boolean _collapsed;
        private Double _height;
        private Boolean _isHidden;
        private Int32 _outlineLevel;

        #endregion Private fields

        #region Constructor

        public XLRow(Int32 row, XLRowParameters xlRowParameters)
            : base(new XLRangeAddress(new XLAddress(xlRowParameters.Worksheet, row, 1, false, false),
                                      new XLAddress(xlRowParameters.Worksheet, row, XLHelper.MaxColumnNumber, false,
                                                    false)),
                  xlRowParameters.IsReference ? xlRowParameters.Worksheet.Internals.RowsCollection[row].StyleValue
                                              : (xlRowParameters.DefaultStyle as XLStyle).Value
                  )
        {
            SetRowNumber(row);

            IsReference = xlRowParameters.IsReference;
            if (IsReference)
                SubscribeToShiftedRows((range, rowShifted) => this.WorksheetRangeShiftedRows(range, rowShifted));
            else
                _height = xlRowParameters.Worksheet.RowHeight;
        }

        public XLRow(XLRow row)
            : base(new XLRangeAddress(new XLAddress(row.Worksheet, row.RowNumber(), 1, false, false),
                                      new XLAddress(row.Worksheet, row.RowNumber(), XLHelper.MaxColumnNumber, false,
                                                    false)),
                  row.StyleValue)
        {
            _height = row._height;
            IsReference = row.IsReference;
            if (IsReference)
                SubscribeToShiftedRows((range, rowShifted) => this.WorksheetRangeShiftedRows(range, rowShifted));

            _collapsed = row._collapsed;
            _isHidden = row._isHidden;
            _outlineLevel = row._outlineLevel;
            HeightChanged = row.HeightChanged;
        }

        #endregion Constructor

        public Boolean IsReference { get; private set; }

        public override IEnumerable<IXLStyle> Styles
        {
            get
            {
                if (IsReference)
                    yield return Worksheet.Internals.RowsCollection[RowNumber()].Style;
                else
                    yield return Style;

                int row = RowNumber();

                foreach (XLCell cell in Worksheet.Internals.CellsCollection.GetCellsInRow(row))
                    yield return cell.Style;
            }
        }

        protected override IEnumerable<XLStylizedBase> Children
        {
            get
            {
                int row = RowNumber();
                if (IsReference)
                    yield return Worksheet.Internals.RowsCollection[row];
                else
                {
                    foreach (XLCell cell in Worksheet.Internals.CellsCollection.GetCellsInRow(row))
                        yield return cell;
                }
            }
        }

        public Boolean Collapsed
        {
            get { return IsReference ? Worksheet.Internals.RowsCollection[RowNumber()].Collapsed : _collapsed; }
            set
            {
                if (IsReference)
                    Worksheet.Internals.RowsCollection[RowNumber()].Collapsed = value;
                else
                    _collapsed = value;
            }
        }

        #region IXLRow Members

        private Boolean _loading;

        public Boolean Loading
        {
            get { return IsReference ? Worksheet.Internals.RowsCollection[RowNumber()].Loading : _loading; }
            set
            {
                if (IsReference)
                    Worksheet.Internals.RowsCollection[RowNumber()].Loading = value;
                else
                    _loading = value;
            }
        }

        public Boolean HeightChanged { get; private set; }

        public Double Height
        {
            get { return IsReference ? Worksheet.Internals.RowsCollection[RowNumber()].Height : _height; }
            set
            {
                if (!Loading)
                    HeightChanged = true;

                if (IsReference)
                    Worksheet.Internals.RowsCollection[RowNumber()].Height = value;
                else
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
            int rowNumber = RowNumber();
            using (var asRange = AsRange())
                asRange.Delete(XLShiftDeletedCells.ShiftCellsUp);

            Worksheet.Internals.RowsCollection.Remove(rowNumber);
            var rowsToMove = new List<Int32>();
            rowsToMove.AddRange(Worksheet.Internals.RowsCollection.Where(c => c.Key > rowNumber).Select(c => c.Key));
            foreach (int row in rowsToMove.OrderBy(r => r))
            {
                Worksheet.Internals.RowsCollection.Add(row - 1, Worksheet.Internals.RowsCollection[row]);
                Worksheet.Internals.RowsCollection.Remove(row);
            }
        }

        public new IXLRows InsertRowsBelow(Int32 numberOfRows)
        {
            int rowNum = RowNumber();
            Worksheet.Internals.RowsCollection.ShiftRowsDown(rowNum + 1, numberOfRows);
            using (var row = Worksheet.Row(rowNum))
            {
                using (var asRange = row.AsRange())
                {
                    asRange.InsertRowsBelowVoid(true, numberOfRows);
                }
            }
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
                internalRow._collapsed = Collapsed;
                internalRow._isHidden = IsHidden;
                internalRow._outlineLevel = OutlineLevel;
            }
        }

        public new IXLRows InsertRowsAbove(Int32 numberOfRows)
        {
            int rowNum = RowNumber();
            if (rowNum > 1)
            {
                using (var row = Worksheet.Row(rowNum - 1))
                {
                    return row.InsertRowsBelow(numberOfRows);
                }
            }

            Worksheet.Internals.RowsCollection.ShiftRowsDown(rowNum, numberOfRows);
            using (var row = Worksheet.Row(rowNum))
            {
                using (var asRange = row.AsRange())
                {
                    asRange.InsertRowsAboveVoid(true, numberOfRows);
                }
            }

            return Worksheet.Rows(rowNum, rowNum + numberOfRows - 1);
        }

        public new IXLRow Clear(XLClearOptions clearOptions = XLClearOptions.All)
        {
            base.Clear(clearOptions);
            return this;
        }

        public IXLCell Cell(Int32 columnNumber)
        {
            return Cell(1, columnNumber);
        }

        public new IXLCell Cell(String columnLetter)
        {
            return Cell(1, columnLetter);
        }

        public new IXLCells Cells()
        {
            return Cells(true, true);
        }

        public new IXLCells Cells(Boolean usedCellsOnly)
        {
            if (usedCellsOnly)
                return Cells(true, true);
            else
                return Cells(FirstCellUsed().Address.ColumnNumber, LastCellUsed().Address.ColumnNumber);
        }

        public new IXLCells Cells(String cellsInRow)
        {
            var retVal = new XLCells(false, false);
            var rangePairs = cellsInRow.Split(',');
            foreach (string pair in rangePairs)
                retVal.Add(Range(pair.Trim()).RangeAddress);
            return retVal;
        }

        public IXLCells Cells(Int32 firstColumn, Int32 lastColumn)
        {
            return Cells(firstColumn + ":" + lastColumn);
        }

        public IXLCells Cells(String firstColumn, String lastColumn)
        {
            return Cells(XLHelper.GetColumnNumberFromLetter(firstColumn) + ":"
                         + XLHelper.GetColumnNumberFromLetter(lastColumn));
        }

        public IXLRow AdjustToContents(Int32 startColumn)
        {
            return AdjustToContents(startColumn, XLHelper.MaxColumnNumber);
        }

        public IXLRow AdjustToContents(Int32 startColumn, Int32 endColumn)
        {
            return AdjustToContents(startColumn, endColumn, 0, Double.MaxValue);
        }

        public IXLRow AdjustToContents(Double minHeight, Double maxHeight)
        {
            return AdjustToContents(1, XLHelper.MaxColumnNumber, minHeight, maxHeight);
        }

        public IXLRow AdjustToContents(Int32 startColumn, Double minHeight, Double maxHeight)
        {
            return AdjustToContents(startColumn, XLHelper.MaxColumnNumber, minHeight, maxHeight);
        }

        public IXLRow AdjustToContents(Int32 startColumn, Int32 endColumn, Double minHeight, Double maxHeight)
        {
            var fontCache = new Dictionary<IXLFontBase, Font>();

            Double rowMaxHeight = minHeight;
            foreach (XLCell c in from XLCell c in Row(startColumn, endColumn).CellsUsed() where !c.IsMerged() select c)
            {
                Double thisHeight;
                Int32 textRotation = c.StyleValue.Alignment.TextRotation;
                if (c.HasRichText || textRotation != 0 || c.InnerText.Contains(Environment.NewLine))
                {
                    var kpList = new List<KeyValuePair<IXLFontBase, string>>();
                    if (c.HasRichText)
                    {
                        foreach (IXLRichString rt in c.RichText)
                        {
                            String formattedString = rt.Text;
                            var arr = formattedString.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
                            Int32 arrCount = arr.Count();
                            for (Int32 i = 0; i < arrCount; i++)
                            {
                                String s = arr[i];
                                if (i < arrCount - 1)
                                    s += Environment.NewLine;
                                kpList.Add(new KeyValuePair<IXLFontBase, String>(rt, s));
                            }
                        }
                    }
                    else
                    {
                        String formattedString = c.GetFormattedString();
                        var arr = formattedString.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
                        Int32 arrCount = arr.Count();
                        for (Int32 i = 0; i < arrCount; i++)
                        {
                            String s = arr[i];
                            if (i < arrCount - 1)
                                s += Environment.NewLine;
                            kpList.Add(new KeyValuePair<IXLFontBase, String>(c.Style.Font, s));
                        }
                    }

                    Double maxLongCol = kpList.Max(kp => kp.Value.Length);
                    Double maxHeightCol = kpList.Max(kp => kp.Key.GetHeight(fontCache));
                    Int32 lineCount = kpList.Count(kp => kp.Value.Contains(Environment.NewLine)) + 1;
                    if (textRotation == 0)
                        thisHeight = maxHeightCol * lineCount;
                    else
                    {
                        if (textRotation == 255)
                            thisHeight = maxLongCol * maxHeightCol;
                        else
                        {
                            Double rotation;
                            if (textRotation == 90 || textRotation == 180)
                                rotation = 90;
                            else
                                rotation = textRotation % 90;

                            thisHeight = (rotation / 90.0) * maxHeightCol * maxLongCol * 0.5;
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

        public Boolean IsHidden
        {
            get { return IsReference ? Worksheet.Internals.RowsCollection[RowNumber()].IsHidden : _isHidden; }
            set
            {
                if (IsReference)
                    Worksheet.Internals.RowsCollection[RowNumber()].IsHidden = value;
                else
                    _isHidden = value;
            }
        }


        public Int32 OutlineLevel
        {
            get { return IsReference ? Worksheet.Internals.RowsCollection[RowNumber()].OutlineLevel : _outlineLevel; }
            set
            {
                if (value < 0 || value > 8)
                    throw new ArgumentOutOfRangeException("value", "Outline level must be between 0 and 8.");

                if (IsReference)
                    Worksheet.Internals.RowsCollection[RowNumber()].OutlineLevel = value;
                else
                {
                    Worksheet.IncrementColumnOutline(value);
                    Worksheet.DecrementColumnOutline(_outlineLevel);
                    _outlineLevel = value;
                }
            }
        }

        public IXLRow Group()
        {
            return Group(false);
        }

        public IXLRow Group(Int32 outlineLevel)
        {
            return Group(outlineLevel, false);
        }

        public IXLRow Ungroup()
        {
            return Ungroup(false);
        }

        public IXLRow Group(Boolean collapse)
        {
            if (OutlineLevel < 8)
                OutlineLevel += 1;

            Collapsed = collapse;
            return this;
        }

        public IXLRow Group(Int32 outlineLevel, Boolean collapse)
        {
            OutlineLevel = outlineLevel;
            Collapsed = collapse;
            return this;
        }

        public IXLRow Ungroup(Boolean ungroupFromAll)
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

        public Int32 CellCount()
        {
            return RangeAddress.LastAddress.ColumnNumber - RangeAddress.FirstAddress.ColumnNumber + 1;
        }

        public new IXLRow Sort()
        {
            return SortLeftToRight();
        }

        public new IXLRow SortLeftToRight(XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false,
                                          Boolean ignoreBlanks = true)
        {
            base.SortLeftToRight(sortOrder, matchCase, ignoreBlanks);
            return this;
        }

        IXLRangeRow IXLRow.CopyTo(IXLCell target)
        {
            using (var asRange = AsRange())
            using (var copy = asRange.CopyTo(target))
                return copy.Row(1);
        }

        IXLRangeRow IXLRow.CopyTo(IXLRangeBase target)
        {
            using (var asRange = AsRange())
            using (var copy = asRange.CopyTo(target))
                return copy.Row(1);
        }

        public IXLRow CopyTo(IXLRow row)
        {
            row.Clear();
            var newRow = (XLRow)row;
            newRow._height = _height;
            newRow.InnerStyle = GetStyle();

            using (var asRange = AsRange())
                asRange.CopyTo(row).Dispose();

            return newRow;
        }

        public IXLRangeRow Row(Int32 start, Int32 end)
        {
            return Range(1, start, 1, end).Row(1);
        }

        public IXLRangeRow Row(IXLCell start, IXLCell end)
        {
            return Row(start.Address.ColumnNumber, end.Address.ColumnNumber);
        }

        public IXLRangeRows Rows(String rows)
        {
            var retVal = new XLRangeRows();
            var rowPairs = rows.Split(',');
            foreach (string pair in rowPairs)
                using (var asRange = AsRange())
                    asRange.Rows(pair.Trim()).ForEach(retVal.Add);
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

        public IXLRangeRow RowUsed(Boolean includeFormats = false)
        {
            return Row(FirstCellUsed(includeFormats), LastCellUsed(includeFormats));
        }

        #endregion IXLRow Members

        public override XLRange AsRange()
        {
            return Range(1, 1, 1, XLHelper.MaxColumnNumber);
        }

        private void WorksheetRangeShiftedRows(XLRange range, int rowsShifted)
        {
            if (range.RangeAddress.IsValid &&
                RangeAddress.IsValid &&
                range.RangeAddress.FirstAddress.RowNumber <= RowNumber())
                SetRowNumber(RowNumber() + rowsShifted);
        }

        private void SetRowNumber(Int32 row)
        {
            if (row <= 0)
                RangeAddress.IsValid = false;
            else
            {
                RangeAddress.IsValid = true;
                RangeAddress.FirstAddress = new XLAddress(Worksheet, row, 1, RangeAddress.FirstAddress.FixedRow,
                                                          RangeAddress.FirstAddress.FixedColumn);
                RangeAddress.LastAddress = new XLAddress(Worksheet,
                                                         row,
                                                         XLHelper.MaxColumnNumber,
                                                         RangeAddress.LastAddress.FixedRow,
                                                         RangeAddress.LastAddress.FixedColumn);
            }
        }

        public override XLRange Range(String rangeAddressStr)
        {
            String rangeAddressToUse;
            if (rangeAddressStr.Contains(':') || rangeAddressStr.Contains('-'))
            {
                if (rangeAddressStr.Contains('-'))
                    rangeAddressStr = rangeAddressStr.Replace('-', ':');

                var arrRange = rangeAddressStr.Split(':');
                string firstPart = arrRange[0];
                string secondPart = arrRange[1];
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
            if (IsReference)
                Worksheet.Internals.RowsCollection[RowNumber()].SetStyleNoColumns(value);
            else
            {
                InnerStyle = value;

                int row = RowNumber();
                foreach (XLCell c in Worksheet.Internals.CellsCollection.GetCellsInRow(row))
                    c.InnerStyle = value;
            }
        }

        private XLRow RowShift(Int32 rowsToShift)
        {
            return Worksheet.Row(RowNumber() + rowsToShift);
        }

        #region XLRow Above

        IXLRow IXLRow.RowAbove()
        {
            return RowAbove();
        }

        IXLRow IXLRow.RowAbove(Int32 step)
        {
            return RowAbove(step);
        }

        public XLRow RowAbove()
        {
            return RowAbove(1);
        }

        public XLRow RowAbove(Int32 step)
        {
            return RowShift(step * -1);
        }

        #endregion XLRow Above

        #region XLRow Below

        IXLRow IXLRow.RowBelow()
        {
            return RowBelow();
        }

        IXLRow IXLRow.RowBelow(Int32 step)
        {
            return RowBelow(step);
        }

        public XLRow RowBelow()
        {
            return RowBelow(1);
        }

        public XLRow RowBelow(Int32 step)
        {
            return RowShift(step);
        }

        #endregion XLRow Below

        public override Boolean IsEmpty()
        {
            return IsEmpty(false);
        }

        public override Boolean IsEmpty(Boolean includeFormats)
        {
            if (includeFormats && !StyleValue.Equals(Worksheet.StyleValue))
                return false;

            return base.IsEmpty(includeFormats);
        }

        public override Boolean IsEntireRow()
        {
            return true;
        }

        public override Boolean IsEntireColumn()
        {
            return false;
        }
    }
}
