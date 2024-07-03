using ClosedXML.Graphics;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal sealed class XLRow : XLRangeBase, IXLRow
    {
        #region Private fields

        /// <summary>
        /// Don't use directly, use properties.
        /// </summary>
        private XlRowFlags _flags;
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

        protected override IEnumerable<XLStylizedBase> Children
        {
            get
            {
                int row = RowNumber();

                foreach (XLCell cell in Worksheet.Internals.CellsCollection.GetCellsInRow(row))
                    yield return cell;
            }
        }

        public Boolean Collapsed
        {
            get => _flags.HasFlag(XlRowFlags.Collapsed);
            set
            {
                if (value)
                    _flags |= XlRowFlags.Collapsed;
                else
                    _flags &= ~XlRowFlags.Collapsed;
            }
        }

        /// <summary>
        /// Distance in pixels from the bottom of the cells in the current row to the typographical
        /// baseline of the cell content if, hypothetically, the zoom level for the sheet containing
        /// this row is 100 percent and the cell has bottom-alignment formatting.
        /// </summary>
        /// <remarks>
        /// If the attribute is set, it sets customHeight to true even if the customHeight is explicitly
        /// set to false. Custom height means no auto-sizing by Excel on load, so if row has this
        /// attribute, it stops Excel from auto-sizing the height of a row to fit the content on load.
        /// </remarks>
        public Double? DyDescent { get; set; }

        /// <summary>
        /// Should cells in the row display phonetic? This doesn't actually affect whether the phonetic are
        /// shown in the row, that depends entirely on the <see cref="IXLCell.ShowPhonetic"/> property
        /// of a cell. This property determines whether a new cell in the row will have it's phonetic turned on
        /// (and also the state of the "Show or hide phonetic" in Excel when whole row is selected).
        /// Default is <c>false</c>.
        /// </summary>
        public Boolean ShowPhonetic
        {
            get => _flags.HasFlag(XlRowFlags.ShowPhonetic);
            set
            {
                if (value)
                    _flags |= XlRowFlags.ShowPhonetic;
                else
                    _flags &= ~XlRowFlags.ShowPhonetic;
            }
        }

        public Boolean Loading
        {
            get => _flags.HasFlag(XlRowFlags.Loading);
            set
            {
                if (value)
                    _flags |= XlRowFlags.Loading;
                else
                    _flags &= ~XlRowFlags.Loading;
            }
        }

        /// <summary>
        /// Does row have an individual height or is it derived from the worksheet <see cref="XLWorksheet.RowHeight"/>?
        /// </summary>
        public Boolean HeightChanged
        {
            get => _flags.HasFlag(XlRowFlags.HeightChanged);
            private set
            {
                if (value)
                    _flags |= XlRowFlags.HeightChanged;
                else
                    _flags &= ~XlRowFlags.HeightChanged;
            }
        }

        #region IXLRow Members

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

        IXLCells IXLRow.Cells(String cellsInRow) => Cells(cellsInRow);

        IXLCells IXLRow.Cells(Int32 firstColumn, Int32 lastColumn) => Cells(firstColumn, lastColumn);

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

        public override XLCells Cells(Boolean usedCellsOnly)
        {
            if (usedCellsOnly)
                return Cells(true, XLCellsUsedOptions.AllContents);
            else
                return Cells(FirstCellUsed().Address.ColumnNumber, LastCellUsed().Address.ColumnNumber);
        }

        public override XLCells Cells(String cellsInRow)
        {
            var retVal = new XLCells(false, XLCellsUsedOptions.AllContents);
            var rangePairs = cellsInRow.Split(',');
            foreach (string pair in rangePairs)
                retVal.Add(Range(pair.Trim()).RangeAddress);
            return retVal;
        }

        public XLCells Cells(Int32 firstColumn, Int32 lastColumn)
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

        public IXLRow AdjustToContents(Int32 startColumn, Int32 endColumn, Double minHeightPt, Double maxHeightPt)
        {
            var engine = Worksheet.Workbook.GraphicEngine;
            var dpi = new Dpi(Worksheet.Workbook.DpiX, Worksheet.Workbook.DpiY);

            var rowHeightPx = CalculateMinRowHeight(startColumn, endColumn, engine, dpi);

            var rowHeightPt = XLHelper.PixelsToPoints(rowHeightPx, dpi.Y);
            if (rowHeightPt <= 0)
                rowHeightPt = Worksheet.RowHeight;

            if (minHeightPt > rowHeightPt)
                rowHeightPt = minHeightPt;

            if (maxHeightPt < rowHeightPt)
                rowHeightPt = maxHeightPt;

            Height = rowHeightPt;

            return this;
        }

        private int CalculateMinRowHeight(int startColumn, int endColumn, IXLGraphicEngine engine, Dpi dpi)
        {
            var glyphs = new List<GlyphBox>();
            XLStyle? cellStyle = null;
            var rowHeightPx = 0;
            foreach (var cell in Row(startColumn, endColumn).CellsUsed().Cast<XLCell>())
            {
                // Clear maintains capacity -> reduce need for GC
                glyphs.Clear();

                if (cell.IsMerged())
                    continue;

                // Reuse styles if possible to reduce memory consumption
                if (cellStyle is null || cellStyle.Value != cell.StyleValue)
                    cellStyle = (XLStyle)cell.Style;

                cell.GetGlyphBoxes(engine, dpi, glyphs);
                var cellHeightPx = (int)Math.Ceiling(GetContentHeight(cellStyle.Alignment.TextRotation, glyphs));

                rowHeightPx = Math.Max(cellHeightPx, rowHeightPx);
            }

            return rowHeightPx;
        }

        private static double GetContentHeight(int textRotationDeg, List<GlyphBox> glyphs)
        {
            if (textRotationDeg == 0)
            {
                var textHeight = 0d;
                var lineMaxHeight = 0d;
                foreach (var glyph in glyphs)
                {
                    if (!glyph.IsLineBreak)
                    {
                        var cellHeightPx = glyph.LineHeight;
                        lineMaxHeight = Math.Max(cellHeightPx, lineMaxHeight);
                    }
                    else
                    {
                        // At the end of each line, add height of the line to total height.
                        textHeight += lineMaxHeight;
                        lineMaxHeight = 0d;
                    }
                }

                // If the last line ends without EOL, it must be also counted
                textHeight += lineMaxHeight;

                return textHeight;
            }
            else if (textRotationDeg == 255)
            {
                // Glyphs are vertically aligned.
                var textHeight = glyphs.Sum(static g => g.LineHeight);
                return textHeight;
            }
            else
            {
                // Rotated text
                var width = 0d;
                var height = 0d;
                foreach (var glyph in glyphs)
                {
                    width += glyph.AdvanceWidth;
                    height = Math.Max(glyph.LineHeight, height);
                }

                var projectedWidth = Math.Sin(XLHelper.DegToRad(textRotationDeg)) * width;
                var projectedHeight = Math.Cos(XLHelper.DegToRad(textRotationDeg)) * height;
                return projectedWidth + projectedHeight;
            }
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
            get => _flags.HasFlag(XlRowFlags.IsHidden);
            set
            {
                if (value)
                    _flags |= XlRowFlags.IsHidden;
                else
                    _flags &= ~XlRowFlags.IsHidden;
            }
        }

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
            // rows are shifted by XLRowCollection
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

        /// <summary>
        /// Flag enum to save space, instead of wasting byte for each flag.
        /// </summary>
        [Flags]
        private enum XlRowFlags : byte
        {
            Collapsed = 1,
            IsHidden = 2,
            ShowPhonetic = 4,
            HeightChanged = 8,
            Loading = 16
        }
    }
}
