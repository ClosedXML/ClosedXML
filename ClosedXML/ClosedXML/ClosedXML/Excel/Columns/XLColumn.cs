using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace ClosedXML.Excel
{
    internal class XLColumn: XLRangeBase, IXLColumn
    {
        public XLColumn(Int32 column, XLColumnParameters xlColumnParameters)
            : base(new XLRangeAddress(new XLAddress(xlColumnParameters.Worksheet, 1, column, false, false), new XLAddress(xlColumnParameters.Worksheet, XLWorksheet.MaxNumberOfRows, column, false, false)))
        {
            SetColumnNumber(column);
            
            this.IsReference = xlColumnParameters.IsReference;
            if (IsReference)
            {
                (Worksheet as XLWorksheet).RangeShiftedColumns += new RangeShiftedColumnsDelegate(Worksheet_RangeShiftedColumns);
            }
            else
            {
                this.style = new XLStyle(this, xlColumnParameters.DefaultStyle);
                this.width = xlColumnParameters.Worksheet.ColumnWidth;
            }
        }

        public XLColumn(XLColumn column)
            : base(new XLRangeAddress(new XLAddress(column.Worksheet, 1, column.ColumnNumber(), false, false), new XLAddress(column.Worksheet, XLWorksheet.MaxNumberOfRows, column.ColumnNumber(), false, false)))
        { 
            width = column.width;
            IsReference = column.IsReference;
            collapsed = column.collapsed;
            isHidden = column.isHidden;
            outlineLevel = column.outlineLevel;
            style = new XLStyle(this, column.Style);
        }

        void Worksheet_RangeShiftedColumns(XLRange range, int columnsShifted)
        {
            if (range.RangeAddress.FirstAddress.ColumnNumber <= this.ColumnNumber())
                SetColumnNumber(this.ColumnNumber() + columnsShifted);
        }

        private void SetColumnNumber(Int32 column)
        {
            if (column <= 0)
            {
                RangeAddress.IsInvalid = false;
            }
            else
            {
                RangeAddress.FirstAddress = new XLAddress(Worksheet, 1, column, RangeAddress.FirstAddress.FixedRow, RangeAddress.FirstAddress.FixedColumn);
                RangeAddress.LastAddress = new XLAddress(Worksheet, XLWorksheet.MaxNumberOfRows, column, RangeAddress.LastAddress.FixedRow, RangeAddress.LastAddress.FixedColumn);
            }
        }

        public Boolean IsReference { get; private set; }

        #region IXLColumn Members

        private Double width;
        public Double Width
        {
            get
            {
                if (IsReference)
                {
                    return (Worksheet as XLWorksheet).Internals.ColumnsCollection[this.ColumnNumber()].Width;
                }
                else
                {
                    return width;
                }
            }
            set
            {
                if (IsReference)
                {
                    (Worksheet as XLWorksheet).Internals.ColumnsCollection[this.ColumnNumber()].Width = value;
                }
                else
                {
                    width = value;
                }
            }
        }

        public void Delete()
        {
            var columnNumber = this.ColumnNumber();
            this.AsRange().Delete(XLShiftDeletedCells.ShiftCellsLeft);
            (Worksheet as XLWorksheet).Internals.ColumnsCollection.Remove(columnNumber);
            List<Int32> columnsToMove = new List<Int32>();
            columnsToMove.AddRange((Worksheet as XLWorksheet).Internals.ColumnsCollection.Where(c => c.Key > columnNumber).Select(c => c.Key));
            foreach (var column in columnsToMove.OrderBy(c=>c))
            {
                (Worksheet as XLWorksheet).Internals.ColumnsCollection.Add(column - 1, (Worksheet as XLWorksheet).Internals.ColumnsCollection[column]);
                (Worksheet as XLWorksheet).Internals.ColumnsCollection.Remove(column);
            }
        }

        public new void Clear()
        {
            var range = this.AsRange();
            range.Clear();
            this.Style = Worksheet.Style;
        }

        public IXLCell Cell(Int32 rowNumber)
        {
            return base.Cell(rowNumber, 1);
        }

        public new IXLCells Cells(String cellsInColumn)
        {
            var retVal = new XLCells(false, false, false);
            var rangePairs = cellsInColumn.Split(',');
            foreach (var pair in rangePairs)
            {
                retVal.Add(Range(pair.Trim()).RangeAddress);
            }
            return retVal;
        }

        public IXLCells Cells(Int32 firstRow, Int32 lastRow)
        {
            return Cells(firstRow + ":" + lastRow);
        }

        #endregion

        #region IXLStylized Members

        private IXLStyle style;
        public override IXLStyle Style
        {
            get
            {
                if (IsReference)
                    return (Worksheet as XLWorksheet).Internals.ColumnsCollection[this.ColumnNumber()].Style;
                else
                    return style;
            }
            set
            {
                if (IsReference)
                {
                    (Worksheet as XLWorksheet).Internals.ColumnsCollection[this.ColumnNumber()].Style = value;
                }
                else
                {
                    style = new XLStyle(this, value);

                    Int32 minRow = 1;
                    Int32 maxRow = 0;
                    var column = this.ColumnNumber();
                    if ((Worksheet as XLWorksheet).Internals.CellsCollection.Values.Any(c => c.Address.ColumnNumber == column))
                    {
                        minRow = (Worksheet as XLWorksheet).Internals.CellsCollection.Values
                            .Where(c => c.Address.ColumnNumber == column)
                            .Min(c => c.Address.RowNumber);
                        maxRow = (Worksheet as XLWorksheet).Internals.CellsCollection.Values
                            .Where(c => c.Address.ColumnNumber == column)
                            .Max(c => c.Address.RowNumber);
                    }

                    if ((Worksheet as XLWorksheet).Internals.RowsCollection.Count > 0)
                    {
                        Int32 minInCollection = (Worksheet as XLWorksheet).Internals.RowsCollection.Keys.Min();
                        Int32 maxInCollection = (Worksheet as XLWorksheet).Internals.RowsCollection.Keys.Max();
                        if (minInCollection < minRow) minRow = minInCollection;
                        if (maxInCollection > maxRow) maxRow = maxInCollection;
                    }

                    for (Int32 ro = minRow; ro <= maxRow; ro++)
                    {
                        Worksheet.Cell(ro, column).Style = value;
                    }
                }
            }
        }

        public override IEnumerable<IXLStyle> Styles
        {
            get
            {
                UpdatingStyle = true;

                yield return Style;

                var column = this.ColumnNumber();
                Int32 minRow = 1;
                Int32 maxRow = 0;
                if ((Worksheet as XLWorksheet).Internals.CellsCollection.Values.Any(c => c.Address.ColumnNumber == column))
                    maxRow = (Worksheet as XLWorksheet).Internals.CellsCollection.Values.Where(c => c.Address.ColumnNumber == column).Max(c => c.Address.RowNumber);
                
                if ((Worksheet as XLWorksheet).Internals.RowsCollection.Count > 0)
                {
                    Int32 maxInCollection = (Worksheet as XLWorksheet).Internals.RowsCollection.Keys.Max();
                    if (maxInCollection > maxRow) maxRow = maxInCollection;
                }

                for (var ro = minRow; ro <= maxRow; ro++)
                {
                    yield return Worksheet.Cell(ro, column).Style;
                }

                UpdatingStyle = false;
            }
        }

        public override Boolean UpdatingStyle { get; set; }

        public override IXLStyle InnerStyle
        {
            get
            {
                if (IsReference)
                    return (Worksheet as XLWorksheet).Internals.ColumnsCollection[this.ColumnNumber()].InnerStyle;
                else
                    return new XLStyle(new XLStylizedContainer(this.style, this), style);
            }
            set
            {
                if (IsReference)
                {
                    (Worksheet as XLWorksheet).Internals.ColumnsCollection[this.ColumnNumber()].InnerStyle = value;
                }
                else
                {
                    style = new XLStyle(this, value);
                }
            }
        }

        #endregion

        public new IXLColumns InsertColumnsAfter(Int32 numberOfColumns)
        {
            var columnNum = this.ColumnNumber();
            (Worksheet as XLWorksheet).Internals.ColumnsCollection.ShiftColumnsRight(columnNum + 1, numberOfColumns);
            XLRange range = (XLRange)this.Worksheet.Column(columnNum).AsRange();
            range.InsertColumnsAfter(true, numberOfColumns);
            return Worksheet.Columns(columnNum + 1, columnNum + numberOfColumns);
        }
        public new IXLColumns InsertColumnsBefore(Int32 numberOfColumns)
        {
            var columnNum = this.ColumnNumber();
            (Worksheet as XLWorksheet).Internals.ColumnsCollection.ShiftColumnsRight(columnNum, numberOfColumns);
            // We can't use this.AsRange() because we've shifted the columns
            // and we want to use the old columnNum.
            XLRange range = (XLRange)this.Worksheet.Column(columnNum).AsRange(); 
            range.InsertColumnsBefore(true, numberOfColumns);
            return Worksheet.Columns(columnNum, columnNum + numberOfColumns - 1);
        }

        public override IXLRange AsRange()
        {
            return Range(1, 1, XLWorksheet.MaxNumberOfRows, 1);
        }
        public override IXLRange Range(String rangeAddressStr)
        {
            String rangeAddressToUse;
            if (rangeAddressStr.Contains(':') || rangeAddressStr.Contains('-'))
            {
                if (rangeAddressStr.Contains('-'))
                    rangeAddressStr = rangeAddressStr.Replace('-', ':');

                String[] arrRange = rangeAddressStr.Split(':');
                var firstPart = arrRange[0];
                var secondPart = arrRange[1];
                rangeAddressToUse = FixColumnAddress(firstPart) + ":" + FixColumnAddress(secondPart);
            }
            else
            {
                rangeAddressToUse = FixColumnAddress(rangeAddressStr);
            }

            var rangeAddress = new XLRangeAddress(Worksheet, rangeAddressToUse);
            return Range(rangeAddress);
        }
        public IXLRangeColumn Range(int firstRow, int lastRow)
        {
            return Range(firstRow, 1, lastRow, 1).Column(1);
        }

        public IXLColumn AdjustToContents()
        {
            return AdjustToContents(1);
        }
        public IXLColumn AdjustToContents(Int32 startRow)
        {
            return AdjustToContents(startRow, XLWorksheet.MaxNumberOfRows);
        }
        public IXLColumn AdjustToContents(Int32 startRow, Int32 endRow)
        {
            Double colMaxWidth = 0;
            foreach (var c in CellsUsed().Where(cell => cell.Address.RowNumber >= startRow && cell.Address.RowNumber <= endRow))
            {
                Boolean isMerged = false;
                var cellAsRange = c.AsRange();
                foreach (var m in (Worksheet as XLWorksheet).Internals.MergedRanges)
                {
                    if (cellAsRange.Intersects(m))
                    {
                        isMerged = true;
                        break;
                    }
                }
                if (!isMerged)
                {
                    var thisWidth = ((XLFont)c.Style.Font).GetWidth(c.GetFormattedString());
                    if (thisWidth > colMaxWidth)
                        colMaxWidth = thisWidth;
                }
            }

            if (colMaxWidth == 0)
                colMaxWidth = Worksheet.ColumnWidth;

            Width = colMaxWidth;

            return this;
        }

        public IXLColumn AdjustToContents(Double minWidth, Double maxWidth)
        {
            return AdjustToContents(1, XLWorksheet.MaxNumberOfRows, minWidth, maxWidth);
        }
        public IXLColumn AdjustToContents(Int32 startRow, Double minWidth, Double maxWidth)
        {
            return AdjustToContents(startRow, XLWorksheet.MaxNumberOfRows, minWidth, maxWidth);
        }
        public IXLColumn AdjustToContents(Int32 startRow, Int32 endRow, Double minWidth, Double maxWidth)
        {
            Double colMaxWidth = minWidth;
            foreach (var c in CellsUsed().Where(cell => cell.Address.RowNumber >= startRow && cell.Address.RowNumber <= endRow))
            {
                Boolean isMerged = false;
                var cellAsRange = c.AsRange();
                foreach (var m in (Worksheet as XLWorksheet).Internals.MergedRanges)
                {
                    if (cellAsRange.Intersects(m))
                    {
                        isMerged = true;
                        break;
                    }
                }
                if (!isMerged)
                {
                    var thisWidth = ((XLFont)c.Style.Font).GetWidth(c.GetFormattedString());
                    if (thisWidth >= maxWidth)
                    {
                        colMaxWidth = maxWidth;
                        break;
                    }
                    else if (thisWidth > colMaxWidth)
                        colMaxWidth = thisWidth;
                }
            }

            Width = colMaxWidth;
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
        private Boolean isHidden;
        public Boolean IsHidden
        {
            get
            {
                if (IsReference)
                {
                    return (Worksheet as XLWorksheet).Internals.ColumnsCollection[this.ColumnNumber()].IsHidden;
                }
                else
                {
                    return isHidden;
                }
            }
            set
            {
                if (IsReference)
                {
                    (Worksheet as XLWorksheet).Internals.ColumnsCollection[this.ColumnNumber()].IsHidden = value;
                }
                else
                {
                    isHidden = value;
                }
            }
        }


        private Boolean collapsed;
        public Boolean Collapsed
        {
            get
            {
                if (IsReference)
                {
                    return (Worksheet as XLWorksheet).Internals.ColumnsCollection[this.ColumnNumber()].Collapsed;
                }
                else
                {
                    return collapsed;
                }
            }
            set
            {
                if (IsReference)
                {
                    (Worksheet as XLWorksheet).Internals.ColumnsCollection[this.ColumnNumber()].Collapsed = value;
                }
                else
                {
                    collapsed = value;
                }
            }
        }
        private Int32 outlineLevel;
        public Int32 OutlineLevel
        {
            get
            {
                if (IsReference)
                {
                    return (Worksheet as XLWorksheet).Internals.ColumnsCollection[this.ColumnNumber()].OutlineLevel;
                }
                else
                {
                    return outlineLevel;
                }
            }
            set
            {
                if (value < 0 || value > 8)
                    throw new ArgumentOutOfRangeException("Outline level must be between 0 and 8.");

                if (IsReference)
                {
                    (Worksheet as XLWorksheet).Internals.ColumnsCollection[this.ColumnNumber()].OutlineLevel = value;
                }
                else
                {
                    (Worksheet as XLWorksheet).IncrementColumnOutline(value);
                    (Worksheet as XLWorksheet).DecrementColumnOutline(outlineLevel);
                    outlineLevel = value;
                }
            }
        }
        public void Group()
        {
            Group(false);
        }
        public void Group(Boolean collapse)
        {
            if (OutlineLevel < 8)
                OutlineLevel += 1;

            Collapsed = collapse;
        }
        public void Group(Int32 outlineLevel)
        {
            Group(outlineLevel, false);
        }
        public void Group(Int32 outlineLevel, Boolean collapse)
        {
            OutlineLevel = outlineLevel;
            Collapsed = collapse;
        }
        public void Ungroup()
        {
            Ungroup(false);
        }
        public void Ungroup(Boolean ungroupFromAll)
        {
            if (ungroupFromAll)
            {
                OutlineLevel = 0;
            }
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
            return this.RangeAddress.LastAddress.ColumnNumber - this.RangeAddress.FirstAddress.ColumnNumber + 1;
        }

        public IXLColumn Sort()
        {
            this.RangeUsed().Sort();
            return this;
        }
        public IXLColumn Sort(XLSortOrder sortOrder)
        {
            this.RangeUsed().Sort(sortOrder);
            return this;
        }
        public IXLColumn Sort(Boolean matchCase)
        {
            this.AsRange().Sort(matchCase);
            return this;
        }
        public IXLColumn Sort(XLSortOrder sortOrder, Boolean matchCase)
        {
            this.AsRange().Sort(sortOrder, matchCase);
            return this;
        }

        private void CopyToCell(IXLRangeColumn rngColumn, IXLCell cell)
        {
            Int32 cellCount = rngColumn.CellCount();
            Int32 roStart = cell.Address.RowNumber;
            Int32 coStart = cell.Address.ColumnNumber;
            for (Int32 ro = roStart; ro <= cellCount + roStart - 1; ro++)
            {
                (cell.Worksheet.Cell(ro, coStart) as XLCell).CopyFrom(rngColumn.Cell(ro - roStart + 1));
            } 
        }

        public new IXLRangeColumn CopyTo(IXLCell target)
        {
            var rngUsed = RangeUsed().Column(1);
            CopyToCell(rngUsed, target);

            Int32 lastRowNumber = target.Address.RowNumber + rngUsed.CellCount() - 1;
            if (lastRowNumber > XLWorksheet.MaxNumberOfRows) lastRowNumber = XLWorksheet.MaxNumberOfRows;

            return target.Worksheet.Range(
                target.Address.RowNumber, 
                target.Address.ColumnNumber,
                lastRowNumber, 
                target.Address.ColumnNumber)
                .Column(1);
        }
        public new IXLRangeColumn CopyTo(IXLRangeBase target)
        {
            var thisRangeUsed = RangeUsed();
            Int32 thisRowCount = thisRangeUsed.RowCount();
            var targetRangeUsed = target.AsRange().RangeUsed();
            Int32 targetRowCount = targetRangeUsed.RowCount();
            Int32 maxRow = thisRowCount > targetRowCount ? thisRowCount : targetRowCount;

            CopyToCell(this.Range(1, 1, maxRow, 1).Column(1), target.FirstCell());

            Int32 lastRowNumber = target.RangeAddress.FirstAddress.RowNumber + maxRow - 1;
            if (lastRowNumber > XLWorksheet.MaxNumberOfRows) lastRowNumber = XLWorksheet.MaxNumberOfRows;

            return (target as XLRangeBase).Worksheet.Range(
                target.RangeAddress.FirstAddress.RowNumber,
                target.RangeAddress.LastAddress.ColumnNumber,
                lastRowNumber,
                target.RangeAddress.LastAddress.ColumnNumber )
                .Column(1);
        }
        public IXLColumn CopyTo(IXLColumn column)
        {
            var thisRangeUsed = RangeUsed();
            Int32 thisRowCount = thisRangeUsed.RowCount();
            //var targetRangeUsed = column target.AsRange().RangeUsed();
            Int32 targetRowCount = column.LastCellUsed(true).Address.RowNumber;
            Int32 maxRow = thisRowCount > targetRowCount ? thisRowCount : targetRowCount;

            CopyToCell(this.Column(1, maxRow), column.FirstCell());
            var newColumn = column as XLColumn;
            newColumn.width = width;
            newColumn.style = new XLStyle(newColumn, Style);
            return newColumn;
        }

        public IXLRangeColumn Column(Int32 start, Int32 end)
        {
            return Range(start, 1, end, 1).Column(1);
        }
        public IXLRangeColumns Columns(String columns)
        {
            var retVal = new XLRangeColumns();
            var columnPairs = columns.Split(',');
            foreach (var pair in columnPairs)
            {
                this.AsRange().Columns(pair.Trim()).ForEach(c => retVal.Add(c));
            }
            return retVal;
        }

    }
}
