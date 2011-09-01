using System;
using System.Collections.Generic;
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


        #endregion

        #region Constructor

        public XLRow(Int32 row, XLRowParameters xlRowParameters)
            : base(new XLRangeAddress(new XLAddress(xlRowParameters.Worksheet, row, 1, false, false),
                                      new XLAddress(xlRowParameters.Worksheet, row, ExcelHelper.MaxColumnNumber, false,
                                                    false)))
        {
            SetRowNumber(row);

            IsReference = xlRowParameters.IsReference;
            if (IsReference)
            {
                //SMELL: Leak may occur
                Worksheet.RangeShiftedRows += WorksheetRangeShiftedRows;
            }
            else
            {
                //_style = new XLStyle(this, xlRowParameters.DefaultStyle);
                SetStyle(xlRowParameters.DefaultStyleId);
                _height = xlRowParameters.Worksheet.RowHeight;
            }
        }

        public XLRow(XLRow row)
            : base(new XLRangeAddress(new XLAddress(row.Worksheet, row.RowNumber(), 1, false, false),
                                      new XLAddress(row.Worksheet, row.RowNumber(), ExcelHelper.MaxColumnNumber, false,
                                                    false)))
        {
            _height = row._height;
            IsReference = row.IsReference;
            _collapsed = row._collapsed;
            _isHidden = row._isHidden;
            _outlineLevel = row._outlineLevel;
            SetStyle(row.GetStyleId());
        }

        #endregion

        public Boolean IsReference { get; private set; }

        public override IEnumerable<IXLStyle> Styles
        {
            get
            {
                UpdatingStyle = true;

                yield return Style;

                int row = RowNumber();

                foreach (var cell in Worksheet.Internals.CellsCollection.GetCellsInRow(row))
                    yield return cell.Style;

                UpdatingStyle = false;
            }
        }

        public override Boolean UpdatingStyle { get; set; }

        public override IXLStyle InnerStyle
        {
            get
            {
                return IsReference
                           ? Worksheet.Internals.RowsCollection[RowNumber()].InnerStyle
                           : GetStyle();
            }
            set
            {
                if (IsReference)
                    Worksheet.Internals.RowsCollection[RowNumber()].InnerStyle = value;
                else
                    SetStyle(value);
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

        public Double Height
        {
            get { return IsReference ? Worksheet.Internals.RowsCollection[RowNumber()].Height : _height; }
            set
            {
                if (IsReference)
                    Worksheet.Internals.RowsCollection[RowNumber()].Height = value;
                else
                    _height = value;
            }
        }

        public void Delete()
        {
            int rowNumber = RowNumber();
            AsRange().Delete(XLShiftDeletedCells.ShiftCellsUp);
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
            var range = (XLRange)Worksheet.Row(rowNum).AsRange();
            range.InsertRowsBelow(true, numberOfRows);
            return Worksheet.Rows(rowNum + 1, rowNum + numberOfRows);
        }

        public new IXLRows InsertRowsAbove(Int32 numberOfRows)
        {
            int rowNum = RowNumber();
            Worksheet.Internals.RowsCollection.ShiftRowsDown(rowNum, numberOfRows);
            // We can't use this.AsRange() because we've shifted the rows
            // and we want to use the old rowNum.
            var range = (XLRange)Worksheet.Row(rowNum).AsRange();
            range.InsertRowsAbove(true, numberOfRows);
            return Worksheet.Rows(rowNum, rowNum + numberOfRows - 1);
        }

        public new IXLRow Clear(XLClearOptions clearOptions = XLClearOptions.ContentsAndFormats)
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
            return CellsUsed(true);
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
            return Cells(ExcelHelper.GetColumnNumberFromLetter(firstColumn) + ":"
                         + ExcelHelper.GetColumnNumberFromLetter(lastColumn));
        }

        public IXLRow AdjustToContents(Int32 startColumn)
        {
            return AdjustToContents(startColumn, ExcelHelper.MaxColumnNumber);
        }

        public IXLRow AdjustToContents(Int32 startColumn, Int32 endColumn)
        {
            return AdjustToContents(startColumn, endColumn, 0, Double.MaxValue);
        }

        public IXLRow AdjustToContents(Double minHeight, Double maxHeight)
        {
            return AdjustToContents(1, ExcelHelper.MaxColumnNumber, minHeight, maxHeight);
        }

        public IXLRow AdjustToContents(Int32 startColumn, Double minHeight, Double maxHeight)
        {
            return AdjustToContents(startColumn, ExcelHelper.MaxColumnNumber, minHeight, maxHeight);
        }

        public IXLRow AdjustToContents(Int32 startColumn, Int32 endColumn, Double minHeight, Double maxHeight)
        {
            Double rowMaxHeight = minHeight;
            foreach (XLCell c in from XLCell c in Row(startColumn, endColumn).CellsUsed() where !c.IsMerged() select c)
            {
                Double thisHeight;
                Int32 textRotation = c.Style.Alignment.TextRotation;
                if (c.HasRichText || textRotation != 0 || c.InnerText.Contains(Environment.NewLine))
                {
                    var kpList = new List<KeyValuePair<IXLFontBase, string>>();
                    if (c.HasRichText)
                    {
                        foreach (IXLRichString rt in c.RichText)
                        {
                            String formattedString = rt.Text;
                            var arr = formattedString.Split(new[] {Environment.NewLine}, StringSplitOptions.None);
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
                        var arr = formattedString.Split(new[] {Environment.NewLine}, StringSplitOptions.None);
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
                    Double maxHeightCol = kpList.Max(kp => kp.Key.GetHeight());
                    Int32 lineCount = kpList.Count(kp => kp.Value.Contains(Environment.NewLine));
                    if (textRotation == 0)
                        thisHeight = maxHeightCol * lineCount;
                    else
                    {
                        if (textRotation == 255)
                            thisHeight = maxLongCol * maxHeightCol;
                        else
                        {
                            Double rotation;
                            if (textRotation == 90 || textRotation == 180 || textRotation == 255)
                                rotation = 90;
                            else
                                rotation = textRotation % 90;

                            thisHeight = (rotation / 90.0) * maxHeightCol * maxLongCol * 0.5;
                        }
                    }
                }
                else
                    thisHeight = c.Style.Font.GetHeight();

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
            return this;
        }

        public void Hide()
        {
            IsHidden = true;
        }

        public void Unhide()
        {
            IsHidden = false;
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

        public override IXLStyle Style
        {
            get { return IsReference ? Worksheet.Internals.RowsCollection[RowNumber()].Style : GetStyle(); }
            set
            {
                if (IsReference)
                    Worksheet.Internals.RowsCollection[RowNumber()].Style = value;
                else
                {
                    SetStyle(value);

                    Int32 minColumn = 1;
                    Int32 maxColumn = 0;
                    int row = RowNumber();
                    if (Worksheet.Internals.CellsCollection.RowsUsed.ContainsKey(row))
                    {
                        minColumn = Worksheet.Internals.CellsCollection.MinColumnInRow(row);
                        maxColumn = Worksheet.Internals.CellsCollection.MaxColumnInRow(row);
                    }

                    if (Worksheet.Internals.ColumnsCollection.Count > 0)
                    {
                        Int32 minInCollection = Worksheet.Internals.ColumnsCollection.Keys.Min();
                        Int32 maxInCollection = Worksheet.Internals.ColumnsCollection.Keys.Max();
                        if (minInCollection < minColumn)
                            minColumn = minInCollection;
                        if (maxInCollection > maxColumn)
                            maxColumn = maxInCollection;
                    }
                    if (minColumn > 0 && maxColumn > 0)
                    {
                        for (Int32 co = minColumn; co <= maxColumn; co++)
                            Worksheet.Cell(row, co).Style = value;
                    }
                }
            }
        }

        public override IXLRange AsRange()
        {
            return Range(1, 1, 1, ExcelHelper.MaxColumnNumber);
        }

        public Int32 OutlineLevel
        {
            get { return IsReference ? Worksheet.Internals.RowsCollection[RowNumber()].OutlineLevel : _outlineLevel; }
            set
            {
                if (value < 1 || value > 8)
                    throw new ArgumentOutOfRangeException("value", "Outline level must be between 1 and 8.");

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

        public void Group()
        {
            Group(false);
        }

        public void Group(Int32 outlineLevel)
        {
            Group(outlineLevel, false);
        }

        public void Ungroup()
        {
            Ungroup(false);
        }

        public void Group(Boolean collapse)
        {
            if (OutlineLevel < 8)
                OutlineLevel += 1;

            Collapsed = collapse;
        }

        public void Group(Int32 outlineLevel, Boolean collapse)
        {
            OutlineLevel = outlineLevel;
            Collapsed = collapse;
        }

        public void Ungroup(Boolean ungroupFromAll)
        {
            if (ungroupFromAll)
                OutlineLevel = 0;
            else
            {
                if (OutlineLevel > 0)
                    OutlineLevel -= 1;
            }
        }

        public void Collapse()
        {
            Collapsed = true;
            Hide();
        }

        public void Expand()
        {
            Collapsed = false;
            Unhide();
        }

        public Int32 CellCount()
        {
            return RangeAddress.LastAddress.ColumnNumber - RangeAddress.FirstAddress.ColumnNumber + 1;
        }

        public IXLRow Sort()
        {
            RangeUsed().Sort(XLSortOrientation.LeftToRight);
            return this;
        }

        public IXLRow Sort(XLSortOrder sortOrder)
        {
            RangeUsed().Sort(XLSortOrientation.LeftToRight, sortOrder);
            return this;
        }

        public IXLRow Sort(Boolean matchCase)
        {
            AsRange().Sort(XLSortOrientation.LeftToRight, matchCase);
            return this;
        }

        public IXLRow Sort(XLSortOrder sortOrder, bool matchCase)
        {
            AsRange().Sort(XLSortOrientation.LeftToRight, sortOrder, matchCase);
            return this;
        }

        IXLRangeRow IXLRow.CopyTo(IXLCell target)
        {
            return AsRange().CopyTo(target).Row(1);
        }

        IXLRangeRow IXLRow.CopyTo(IXLRangeBase target)
        {
            return AsRange().CopyTo(target).Row(1);
        }

        public IXLRow CopyTo(IXLRow row)
        {
            row.Clear();
            AsRange().CopyTo(row);

            var newRow = (XLRow)row;
            newRow._height = _height;
            //newRow._style = new XLStyle(newRow, Style);
            newRow.Style = GetStyle();

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

        public IXLRow SetDataType(XLCellValues dataType)
        {
            DataType = dataType;
            return this;
        }

        #endregion

        private void WorksheetRangeShiftedRows(XLRange range, int rowsShifted)
        {
            if (range.RangeAddress.FirstAddress.RowNumber <= RowNumber())
                SetRowNumber(RowNumber() + rowsShifted);
        }

        private void SetRowNumber(Int32 row)
        {
            if (row <= 0)
                RangeAddress.IsInvalid = false;
            else
            {
                RangeAddress.FirstAddress = new XLAddress(Worksheet, row, 1, RangeAddress.FirstAddress.FixedRow,
                                                          RangeAddress.FirstAddress.FixedColumn);
                RangeAddress.LastAddress = new XLAddress(Worksheet,
                                                         row,
                                                         ExcelHelper.MaxColumnNumber,
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
                SetStyle(value);

                int row = RowNumber();
                foreach (XLCell c in Worksheet.Internals.CellsCollection.GetCellsInRow(row))
                    c.Style = value;
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

        #endregion

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

        #endregion

        public IXLRangeRow RowUsed(Boolean includeFormats = false)
        {
            return Row(FirstCellUsed(includeFormats), LastCellUsed(includeFormats));
        }
    }
}