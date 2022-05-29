using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLRow : XLRangeBase, IXLRow
    {
        #region Private fields

        private Double _height;
        private Int32 _outlineLevel;

        #endregion Private fields

        #region Constructor

        /// <summary>
        /// The direct constructor should only be used in <see cref="XLWorksheet.RangeFactory"/>.
        /// </summary>
        public XLRow(XLWorksheet worksheet, Int32 row)
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

                foreach (XLCell cell in Worksheet.Internals.CellsCollection.GetCellsInRow(row))
                    yield return cell;
            }
        }

        public Boolean Collapsed { get; set; }

        #region IXLRow Members

        public Boolean Loading { get; set; }

        public Boolean HeightChanged { get; private set; }

        public Double Height
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
            int rowNumber = RowNumber();
            AsRange().Delete(XLShiftDeletedCells.ShiftCellsUp);
            Worksheet.DeleteRow(rowNumber);
        }

        public new IXLRows InsertRowsBelow(Int32 numberOfRows)
        {
            int rowNum = RowNumber();
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

        public new IXLRows InsertRowsAbove(Int32 numberOfRows)
        {
            int rowNum = RowNumber();
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

        public IXLCell Cell(Int32 columnNumber)
        {
            return Cell(1, columnNumber);
        }

        public override XLCell Cell(String columnLetter)
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

        public override IXLCells Cells(Boolean usedCellsOnly)
        {
            if (usedCellsOnly)
                return Cells(true, XLCellsUsedOptions.AllContents);
            else
                return Cells(FirstCellUsed().Address.ColumnNumber, LastCellUsed().Address.ColumnNumber);
        }

        public override IXLCells Cells(String cellsInRow)
        {
            var retVal = new XLCells(false, XLCellsUsedOptions.AllContents);
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
                        foreach (IXLRichString rt in c.GetRichText())
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

        public Boolean IsHidden { get; set; }

        public Int32 OutlineLevel
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
        public IXLRangeRow RowUsed(Boolean includeFormats)
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

        internal void SetRowNumber(Int32 row)
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
            InnerStyle = value;

            int row = RowNumber();
            foreach (XLCell c in Worksheet.Internals.CellsCollection.GetCellsInRow(row))
                c.InnerStyle = value;
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
            return IsEmpty(XLCellsUsedOptions.AllContents);
        }

        public override Boolean IsEmpty(XLCellsUsedOptions options)
        {
            if (options.HasFlag(XLCellsUsedOptions.NormalFormats) &&
                !StyleValue.Equals(Worksheet.StyleValue))
                return false;

            return base.IsEmpty(options);
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
