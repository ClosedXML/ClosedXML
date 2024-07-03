using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Graphics;

namespace ClosedXML.Excel
{
    internal class XLColumn : XLRangeBase, IXLColumn
    {
        #region Private fields

        private int _outlineLevel;

        #endregion Private fields

        #region Constructor

        /// <summary>
        /// The direct constructor should only be used in <see cref="XLWorksheet.RangeFactory"/>.
        /// </summary>
        public XLColumn(XLWorksheet worksheet, Int32 column)
            : base(XLRangeAddress.EntireColumn(worksheet, column), worksheet.StyleValue)
        {
            SetColumnNumber(column);

            Width = worksheet.ColumnWidth;
        }

        #endregion Constructor

        public override XLRangeType RangeType
        {
            get { return XLRangeType.Column; }
        }

        protected override IEnumerable<XLStylizedBase> Children
        {
            get
            {
                int column = ColumnNumber();
                foreach (XLCell cell in Worksheet.Internals.CellsCollection.GetCellsInColumn(column))
                    yield return cell;
            }
        }

        public Boolean Collapsed { get; set; }

        #region IXLColumn Members

        public Double Width { get; set; }

        IXLCells IXLColumn.Cells(String cellsInColumn) => Cells(cellsInColumn);

        IXLCells IXLColumn.Cells(Int32 firstRow, Int32 lastRow) => Cells(firstRow, lastRow);

        public void Delete()
        {
            int columnNumber = ColumnNumber();
            Delete(XLShiftDeletedCells.ShiftCellsLeft);
            Worksheet.DeleteColumn(columnNumber);
        }

        public new IXLColumn Clear(XLClearOptions clearOptions = XLClearOptions.All)
        {
            base.Clear(clearOptions);
            return this;
        }

        public IXLCell Cell(Int32 rowNumber)
        {
            return Cell(rowNumber, 1);
        }

        public override XLCells Cells(String cellsInColumn)
        {
            var retVal = new XLCells(false, XLCellsUsedOptions.All);
            var rangePairs = cellsInColumn.Split(',');
            foreach (string pair in rangePairs)
                retVal.Add(Range(pair.Trim()).RangeAddress);
            return retVal;
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
                return Cells(FirstCellUsed().Address.RowNumber, LastCellUsed().Address.RowNumber);
        }

        public XLCells Cells(Int32 firstRow, Int32 lastRow)
        {
            return Cells(firstRow + ":" + lastRow);
        }

        public new IXLColumns InsertColumnsAfter(Int32 numberOfColumns)
        {
            int columnNum = ColumnNumber();
            Worksheet.Internals.ColumnsCollection.ShiftColumnsRight(columnNum + 1, numberOfColumns);
            Worksheet.Column(columnNum).InsertColumnsAfterVoid(true, numberOfColumns);
            var newColumns = Worksheet.Columns(columnNum + 1, columnNum + numberOfColumns);
            CopyColumns(newColumns);
            return newColumns;
        }

        public new IXLColumns InsertColumnsBefore(Int32 numberOfColumns)
        {
            int columnNum = ColumnNumber();
            if (columnNum > 1)
            {
                return Worksheet.Column(columnNum - 1).InsertColumnsAfter(numberOfColumns);
            }

            Worksheet.Internals.ColumnsCollection.ShiftColumnsRight(columnNum, numberOfColumns);
            Worksheet.Column(columnNum).InsertColumnsBeforeVoid(true, numberOfColumns);

            return Worksheet.Columns(columnNum, columnNum + numberOfColumns - 1);
        }

        private void CopyColumns(IXLColumns newColumns)
        {
            foreach (var newColumn in newColumns)
            {
                var internalColumn = Worksheet.Internals.ColumnsCollection[newColumn.ColumnNumber()];
                internalColumn.Width = Width;
                internalColumn.InnerStyle = InnerStyle;
                internalColumn.Collapsed = Collapsed;
                internalColumn.IsHidden = IsHidden;
                internalColumn._outlineLevel = OutlineLevel;
            }
        }

        public IXLColumn AdjustToContents()
        {
            return AdjustToContents(1);
        }

        public IXLColumn AdjustToContents(Int32 startRow)
        {
            return AdjustToContents(startRow, XLHelper.MaxRowNumber);
        }

        public IXLColumn AdjustToContents(Int32 startRow, Int32 endRow)
        {
            return AdjustToContents(startRow, endRow, 0, Double.MaxValue);
        }

        public IXLColumn AdjustToContents(Double minWidth, Double maxWidth)
        {
            return AdjustToContents(1, XLHelper.MaxRowNumber, minWidth, maxWidth);
        }

        public IXLColumn AdjustToContents(Int32 startRow, Double minWidth, Double maxWidth)
        {
            return AdjustToContents(startRow, XLHelper.MaxRowNumber, minWidth, maxWidth);
        }

        public IXLColumn AdjustToContents(Int32 startRow, Int32 endRow, Double minWidthNoC, Double maxWidthNoC)
        {
            var engine = Worksheet.Workbook.GraphicEngine;
            var dpi = new Dpi(Worksheet.Workbook.DpiX, Worksheet.Workbook.DpiY);
            var columnWidthPx = CalculateMinColumnWidth(startRow, endRow, engine, dpi);

            // Maximum digit width, rounded to pixels, so Calibri at 11 pts returns 7 pixels MDW (the correct value)
            var mdw = (int)Math.Round(engine.GetMaxDigitWidth(Worksheet.Workbook.Style.Font, dpi.X));

            var minWidthInPx = Math.Ceiling(XLHelper.NoCToPixels(minWidthNoC, mdw));
            if (columnWidthPx < minWidthInPx)
                columnWidthPx = (int)minWidthInPx;

            var maxWidthInPx = Math.Ceiling(XLHelper.NoCToPixels(maxWidthNoC, mdw));
            if (columnWidthPx > maxWidthInPx)
                columnWidthPx = (int)maxWidthInPx;

            var colMaxWidth = XLHelper.PixelToNoC(columnWidthPx, mdw);

            // If there is nothing in the column, use worksheet column width.
            if (colMaxWidth <= 0)
                colMaxWidth = Worksheet.ColumnWidth;

            Width = colMaxWidth;

            return this;
        }

        /// <summary>
        /// Calculate column width in pixels according to the content of cells.
        /// </summary>
        /// <param name="startRow">First row number whose content is used for determination.</param>
        /// <param name="endRow">Last row number whose content is used for determination.</param>
        /// <param name="engine">Engine to determine size of glyphs.</param>
        /// <param name="dpi">DPI of the worksheet.</param>
        private int CalculateMinColumnWidth(int startRow, int endRow, IXLGraphicEngine engine, Dpi dpi)
        {
            var autoFilterRows = new List<Int32>();
            if (this.Worksheet.AutoFilter != null && Worksheet.AutoFilter.Range != null)
                autoFilterRows.Add(this.Worksheet.AutoFilter.Range.FirstRow().RowNumber());

            autoFilterRows.AddRange(Worksheet.Tables.Where<XLTable>(t =>
                    t.AutoFilter != null
                    && t.AutoFilter.Range != null
                    && !autoFilterRows.Contains(t.AutoFilter.Range.FirstRow().RowNumber()))
                .Select(t => t.AutoFilter.Range.FirstRow().RowNumber()));

            // Reusable buffer
            var glyphs = new List<GlyphBox>();
            XLStyle? cellStyle = null;
            var columnWidthPx = 0;
            foreach (var cell in Column(startRow, endRow).CellsUsed())
            {
                // Clear maintains capacity -> reduce need for GC
                glyphs.Clear();

                if (cell.IsMerged())
                    continue;

                // Reuse styles if possible to reduce memory consumption
                if (cellStyle is null || cellStyle.Value != cell.StyleValue)
                    cellStyle = (XLStyle)cell.Style;

                cell.GetGlyphBoxes(engine, dpi, glyphs);
                var textWidthPx = (int)Math.Ceiling(GetContentWidth(cellStyle.Alignment.TextRotation, glyphs));

                var scaledMdw = engine.GetMaxDigitWidth(cellStyle.Font, dpi.X);
                scaledMdw = Math.Round(scaledMdw, MidpointRounding.AwayFromZero);

                // Not sure about rounding, but larger is probably better, so use ceiling.
                // Due to mismatched rendering, add 3% instead of 1.75%, to have additional space.
                var oneSidePadding = (int)Math.Ceiling(textWidthPx * 0.03 + scaledMdw / 4);

                // Cell width if calculated as content width + padding on each side of a content.
                // The one side padding is roughly 1.75% of content + MDW/4.
                // The additional pixel is there for lines between cells.
                var cellWidthPx = textWidthPx + 2 * oneSidePadding + 1;

                if (autoFilterRows.Contains(cell.Address.RowNumber))
                {
                    // Autofilter arrow is 16px at 96dpi, scaling through DPI, e.g. 20px at 120dpi
                    cellWidthPx += (int)Math.Round(16d * dpi.X / 96d, MidpointRounding.AwayFromZero);
                }

                columnWidthPx = Math.Max(cellWidthPx, columnWidthPx);
            }

            return columnWidthPx;
        }

        private static double GetContentWidth(int textRotationDeg, List<GlyphBox> glyphs)
        {
            if (textRotationDeg == 0)
            {
                var maxTextWidth = 0d;
                var lineTextWidth = 0d;
                foreach (var glyph in glyphs)
                {
                    if (!glyph.IsLineBreak)
                    {
                        lineTextWidth += glyph.AdvanceWidth;
                        maxTextWidth = Math.Max(lineTextWidth, maxTextWidth);
                    }
                    else
                        lineTextWidth = 0;
                }

                return maxTextWidth;
            }
            if (textRotationDeg == 255)
            {
                // Glyphs are arranged vertically, top to bottom.
                var maxGlyphWidth = 0d;
                foreach (var grapheme in glyphs)
                    maxGlyphWidth = Math.Max(grapheme.AdvanceWidth, maxGlyphWidth);

                return maxGlyphWidth;
            }
            else
            {
                // Glyphs are rotated
                if (textRotationDeg > 90)
                    textRotationDeg = 90 - textRotationDeg;

                var totalWidth = 0d;
                var maxHeight = 0d;
                foreach (var glyph in glyphs)
                {
                    totalWidth += glyph.AdvanceWidth;
                    maxHeight = Math.Max(maxHeight, glyph.LineHeight);
                }

                var projectedHeight = maxHeight * Math.Cos(XLHelper.DegToRad(90 - textRotationDeg));
                var projectedWidth = totalWidth * Math.Cos(XLHelper.DegToRad(textRotationDeg));
                return projectedWidth + projectedHeight;
            }
        }

        public IXLColumn Hide()
        {
            IsHidden = true;
            return this;
        }

        public IXLColumn Unhide()
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

        public IXLColumn Group()
        {
            return Group(false);
        }

        public IXLColumn Group(Boolean collapse)
        {
            if (OutlineLevel < 8)
                OutlineLevel += 1;

            Collapsed = collapse;
            return this;
        }

        public IXLColumn Group(Int32 outlineLevel)
        {
            return Group(outlineLevel, false);
        }

        public IXLColumn Group(Int32 outlineLevel, Boolean collapse)
        {
            OutlineLevel = outlineLevel;
            Collapsed = collapse;
            return this;
        }

        public IXLColumn Ungroup()
        {
            return Ungroup(false);
        }

        public IXLColumn Ungroup(Boolean ungroupFromAll)
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

        public IXLColumn Collapse()
        {
            Collapsed = true;
            return Hide();
        }

        public IXLColumn Expand()
        {
            Collapsed = false;
            return Unhide();
        }

        public Int32 CellCount()
        {
            return RangeAddress.LastAddress.ColumnNumber - RangeAddress.FirstAddress.ColumnNumber + 1;
        }

        public IXLColumn Sort(XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false,
                              Boolean ignoreBlanks = true)
        {
            Sort(1, sortOrder, matchCase, ignoreBlanks);
            return this;
        }

        IXLRangeColumn IXLColumn.Column(Int32 start, Int32 end) => Column(start, end);

        IXLRangeColumn IXLColumn.CopyTo(IXLCell target)
        {
            var copy = AsRange().CopyTo(target);
            return copy.Column(1);
        }

        IXLRangeColumn IXLColumn.CopyTo(IXLRangeBase target)
        {
            var copy = AsRange().CopyTo(target);
            return copy.Column(1);
        }

        public IXLColumn CopyTo(IXLColumn column)
        {
            column.Clear();
            var newColumn = (XLColumn)column;
            newColumn.Width = Width;
            newColumn.InnerStyle = InnerStyle;
            newColumn.IsHidden = IsHidden;

            (this as XLRangeBase).CopyTo(column);

            return newColumn;
        }

        public XLRangeColumn Column(Int32 start, Int32 end)
        {
            return Range(start, 1, end, 1).Column(1);
        }

        public IXLRangeColumn Column(IXLCell start, IXLCell end)
        {
            return Column(start.Address.RowNumber, end.Address.RowNumber);
        }

        public IXLRangeColumns Columns(String columns)
        {
            var retVal = new XLRangeColumns();
            var columnPairs = columns.Split(',');
            foreach (string pair in columnPairs)
                AsRange().Columns(pair.Trim()).ForEach(retVal.Add);
            return retVal;
        }

        /// <summary>
        ///   Adds a vertical page break after this column.
        /// </summary>
        public IXLColumn AddVerticalPageBreak()
        {
            Worksheet.PageSetup.AddVerticalPageBreak(ColumnNumber());
            return this;
        }

        public IXLRangeColumn ColumnUsed(XLCellsUsedOptions options = XLCellsUsedOptions.AllContents)
        {
            return Column((this as IXLRangeBase).FirstCellUsed(options),
                          (this as IXLRangeBase).LastCellUsed(options));
        }

        #endregion IXLColumn Members

        public override XLRange AsRange()
        {
            return Range(1, 1, XLHelper.MaxRowNumber, 1);
        }

        internal override void WorksheetRangeShiftedColumns(XLRange range, int columnsShifted)
        {
            return; // Columns are shifted by XLColumnCollection
        }

        internal override void WorksheetRangeShiftedRows(XLRange range, int rowsShifted)
        {
            //do nothing
        }

        internal void SetColumnNumber(int column)
        {
            RangeAddress = new XLRangeAddress(
                new XLAddress(Worksheet,
                              1,
                              column,
                              RangeAddress.FirstAddress.FixedRow,
                              RangeAddress.FirstAddress.FixedColumn),
                new XLAddress(Worksheet,
                              XLHelper.MaxRowNumber,
                              column,
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
                rangeAddressToUse = FixColumnAddress(firstPart) + ":" + FixColumnAddress(secondPart);
            }
            else
                rangeAddressToUse = FixColumnAddress(rangeAddressStr);

            var rangeAddress = new XLRangeAddress(Worksheet, rangeAddressToUse);
            return Range(rangeAddress);
        }

        public IXLRangeColumn Range(int firstRow, int lastRow)
        {
            return Range(firstRow, 1, lastRow, 1).Column(1);
        }

        private XLColumn ColumnShift(Int32 columnsToShift)
        {
            return Worksheet.Column(ColumnNumber() + columnsToShift);
        }

        #region XLColumn Left

        IXLColumn IXLColumn.ColumnLeft()
        {
            return ColumnLeft();
        }

        IXLColumn IXLColumn.ColumnLeft(Int32 step)
        {
            return ColumnLeft(step);
        }

        public XLColumn ColumnLeft()
        {
            return ColumnLeft(1);
        }

        public XLColumn ColumnLeft(Int32 step)
        {
            return ColumnShift(step * -1);
        }

        #endregion XLColumn Left

        #region XLColumn Right

        IXLColumn IXLColumn.ColumnRight()
        {
            return ColumnRight();
        }

        IXLColumn IXLColumn.ColumnRight(Int32 step)
        {
            return ColumnRight(step);
        }

        public XLColumn ColumnRight()
        {
            return ColumnRight(1);
        }

        public XLColumn ColumnRight(Int32 step)
        {
            return ColumnShift(step);
        }

        #endregion XLColumn Right

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
            return false;
        }

        public override Boolean IsEntireColumn()
        {
            return true;
        }
    }
}
