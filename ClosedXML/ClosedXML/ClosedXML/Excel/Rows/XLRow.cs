using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace ClosedXML.Excel
{
    internal class XLRow: XLRangeBase, IXLRow
    {
        public XLRow(Int32 row, XLRowParameters xlRowParameters)
            : base(new XLRangeAddress(new XLAddress(xlRowParameters.Worksheet, row, 1, false, false), new XLAddress(xlRowParameters.Worksheet, row, XLWorksheet.MaxNumberOfColumns, false, false)))
        {
            SetRowNumber(row);
            
            this.IsReference = xlRowParameters.IsReference;
            if (IsReference)
            {
                (Worksheet as XLWorksheet).RangeShiftedRows += new RangeShiftedRowsDelegate(Worksheet_RangeShiftedRows);
            }
            else
            {
                this.style = new XLStyle(this, xlRowParameters.DefaultStyle);
                this.height = xlRowParameters.Worksheet.RowHeight;
            }
        }

        public XLRow(XLRow row)
            : base(new XLRangeAddress(new XLAddress(row.Worksheet, row.RowNumber(), 1, false, false), new XLAddress(row.Worksheet, row.RowNumber(), XLWorksheet.MaxNumberOfColumns, false, false)))
        {
            height = row.height;
            IsReference = row.IsReference;
            collapsed = row.collapsed;
            isHidden = row.isHidden;
            outlineLevel = row.outlineLevel;
            style = new XLStyle(this, row.Style);
        }

        void Worksheet_RangeShiftedRows(XLRange range, int rowsShifted)
        {
            if (range.RangeAddress.FirstAddress.RowNumber <= this.RowNumber())
                SetRowNumber(this.RowNumber() + rowsShifted);
        }

        void RowsCollection_RowShifted(int startingRow, int rowsShifted)
        {
            if (startingRow <= this.RowNumber())
                SetRowNumber(this.RowNumber() + rowsShifted);
        }

        private void SetRowNumber(Int32 row)
        {
            if (row <= 0)
            {
                RangeAddress.IsInvalid = false;
            }
            else
            {
                RangeAddress.FirstAddress = new XLAddress(Worksheet, row, 1, RangeAddress.FirstAddress.FixedRow, RangeAddress.FirstAddress.FixedColumn);
                RangeAddress.LastAddress = new XLAddress(Worksheet, row, XLWorksheet.MaxNumberOfColumns, RangeAddress.LastAddress.FixedRow, RangeAddress.LastAddress.FixedColumn);
            }
        }

        public Boolean IsReference { get; private set; }

        #region IXLRow Members

        private Double height;
        public Double Height 
        {
            get
            {
                if (IsReference)
                {
                    return (Worksheet as XLWorksheet).Internals.RowsCollection[this.RowNumber()].Height;
                }
                else
                {
                    return height;
                }
            }
            set
            {
                if (IsReference)
                {
                    (Worksheet as XLWorksheet).Internals.RowsCollection[this.RowNumber()].Height = value;
                }
                else
                {
                    height = value;
                }
            }
        }

        public void Delete()
        {
            var rowNumber = this.RowNumber();
            this.AsRange().Delete(XLShiftDeletedCells.ShiftCellsUp);
            (Worksheet as XLWorksheet).Internals.RowsCollection.Remove(rowNumber);
            List<Int32> rowsToMove = new List<Int32>();
            rowsToMove.AddRange((Worksheet as XLWorksheet).Internals.RowsCollection.Where(c => c.Key > rowNumber).Select(c => c.Key));
            foreach (var row in rowsToMove.OrderBy(r=>r))
            {
                (Worksheet as XLWorksheet).Internals.RowsCollection.Add(row - 1, (Worksheet as XLWorksheet).Internals.RowsCollection[row]);
                (Worksheet as XLWorksheet).Internals.RowsCollection.Remove(row);
            }
        }


        public new IXLRows InsertRowsBelow(Int32 numberOfRows)
        {
            var rowNum = this.RowNumber();
            (Worksheet as XLWorksheet).Internals.RowsCollection.ShiftRowsDown(rowNum + 1, numberOfRows);
            XLRange range = (XLRange)this.Worksheet.Row(rowNum).AsRange();
            range.InsertRowsBelow(true, numberOfRows);
            return Worksheet.Rows(rowNum + 1, rowNum + numberOfRows);
        }

        public new IXLRows InsertRowsAbove(Int32 numberOfRows)
        {
            var rowNum = this.RowNumber();
            (Worksheet as XLWorksheet).Internals.RowsCollection.ShiftRowsDown(rowNum, numberOfRows);
            // We can't use this.AsRange() because we've shifted the rows
            // and we want to use the old rowNum.
            XLRange range = (XLRange)this.Worksheet.Row(rowNum).AsRange(); 
            range.InsertRowsAbove(true, numberOfRows);
            return Worksheet.Rows(rowNum, rowNum + numberOfRows - 1);
        }

        public new void Clear()
        {
            var range = this.AsRange();
            range.Clear();
            this.Style = Worksheet.Style;
        }

        public IXLCell Cell(Int32 columnNumber)
        {
            return base.Cell(1, columnNumber);
        }
        public new IXLCell Cell(String columnLetter)
        {
            return base.Cell(1, columnLetter);
        }

        public new IXLCells Cells(String cellsInRow)
        {
            var retVal = new XLCells(false, false, false);
            var rangePairs = cellsInRow.Split(',');
            foreach (var pair in rangePairs)
            {
                retVal.Add(Range(pair.Trim()).RangeAddress);
            }
            return retVal;
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
                rangeAddressToUse = FixRowAddress(firstPart) + ":" + FixRowAddress(secondPart);
            }
            else
            {
                rangeAddressToUse = FixRowAddress(rangeAddressStr);
            }

            var rangeAddress = new XLRangeAddress(Worksheet, rangeAddressToUse);
            return Range(rangeAddress);
        }

        public IXLCells Cells(Int32 firstColumn, Int32 lastColumn)
        {
            return Cells(firstColumn + ":" + lastColumn);
        }

        public IXLCells Cells(String firstColumn, String lastColumn)
        {
            return Cells(XLAddress.GetColumnNumberFromLetter(firstColumn) + ":" 
                + XLAddress.GetColumnNumberFromLetter(lastColumn));
        }
        public IXLRow AdjustToContents()
        {
            return AdjustToContents(1);
        }
        public IXLRow AdjustToContents(Int32 startColumn)
        {
            return AdjustToContents(startColumn, XLWorksheet.MaxNumberOfColumns);
        }
        public IXLRow AdjustToContents(Int32 startColumn, Int32 endColumn)
        {
            Double maxHeight = 0;
            foreach (var c in CellsUsed().Where(cell => cell.Address.ColumnNumber >= startColumn && cell.Address.ColumnNumber <= endColumn))
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
                    //var thisHeight = ((XLFont)c.Style.Font).GetHeight();

                    Int32 textRotation = c.Style.Alignment.TextRotation;
                    var f = (XLFont)c.Style.Font;
                    Double thisHeight;
                    if (textRotation == 0)
                    {
                        thisHeight = f.GetHeight();
                    }
                    else
                    {
                        if (textRotation == 255)
                        {
                            thisHeight = f.GetHeight() * c.GetFormattedString().Length;
                        }
                        else
                        {
                            Int32 rotation;
                            if (textRotation == 90 || textRotation == 180 || textRotation == 255)
                                rotation = 90;
                            else
                                rotation = textRotation % 90;

                            Double r = DegreeToRadian(rotation);
                            Double b = f.GetHeight();
                            Double m = f.GetHeight() * c.GetFormattedString().Length;
                            Double t = m - b;
                            thisHeight = (rotation / 90) * t;
                            
                        }
                    }

                    if (thisHeight > maxHeight)
                        maxHeight = thisHeight;
                }
            }

            if (maxHeight == 0)
                maxHeight = Worksheet.RowHeight;

            Height = maxHeight;
            return this;
        }

        private double DegreeToRadian(double angle)
        {
            return Math.PI * angle / 180.0;
        }

        public IXLRow AdjustToContents(Double minHeight, Double maxHeight)
        {
            return AdjustToContents(1, XLWorksheet.MaxNumberOfColumns, minHeight, maxHeight);
        }
        public IXLRow AdjustToContents(Int32 startColumn, Double minHeight, Double maxHeight)
        {
            return AdjustToContents(startColumn, XLWorksheet.MaxNumberOfColumns, minHeight, maxHeight);
        }
        public IXLRow AdjustToContents(Int32 startColumn, Int32 endColumn, Double minHeight, Double maxHeight)
        {
            Double rowMaxHeight = minHeight;
            foreach (var c in CellsUsed().Where(cell => cell.Address.ColumnNumber >= startColumn && cell.Address.ColumnNumber <= endColumn))
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
                    Int32 textRotation = c.Style.Alignment.TextRotation;
                    var f = (XLFont)c.Style.Font;
                    Double thisHeight;
                    if (textRotation == 0)
                    {
                        thisHeight = f.GetHeight();
                    }
                    else
                    {
                        if (textRotation == 255)
                        {
                            thisHeight = f.GetHeight() * c.GetFormattedString().Length;
                        }
                        else
                        {
                            Int32 rotation;
                            if (textRotation == 90 || textRotation == 180 || textRotation == 255)
                                rotation = 90;
                            else
                                rotation = textRotation % 90;

                            Double r = DegreeToRadian(rotation);
                            Double b = f.GetHeight();
                            Double m = f.GetHeight() * c.GetFormattedString().Length;
                            Double t = m - b;
                            thisHeight = (rotation / 90) * t;

                        }
                    }

                    if (thisHeight >= maxHeight)
                    {
                        rowMaxHeight = maxHeight;
                        break;
                    }
                    else if (thisHeight > rowMaxHeight)
                        rowMaxHeight = thisHeight;
                }
            }

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
        private Boolean isHidden;
        public Boolean IsHidden
        {
            get
            {
                if (IsReference)
                {
                    return (Worksheet as XLWorksheet).Internals.RowsCollection[this.RowNumber()].IsHidden;
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
                    (Worksheet as XLWorksheet).Internals.RowsCollection[this.RowNumber()].IsHidden = value;
                }
                else
                {
                    isHidden = value;
                }
            }
        }

        #endregion

        #region IXLStylized Members

        internal void SetStyleNoColumns(IXLStyle value)
        {
            if (IsReference)
            {
                (Worksheet as XLWorksheet).Internals.RowsCollection[this.RowNumber()].SetStyleNoColumns(value);
            }
            else
            {
                style = new XLStyle(this, value);

                var row = this.RowNumber();
                foreach (var c in (Worksheet as XLWorksheet).Internals.CellsCollection.Values.Where(c => c.Address.RowNumber == row))
                {
                    c.Style = value;
                }
            }
        }

        internal IXLStyle style;
        public override IXLStyle Style
        {
            get
            {
                if (IsReference)
                    return (Worksheet as XLWorksheet).Internals.RowsCollection[this.RowNumber()].Style;
                else
                    return style;
            }
            set
            {
                if (IsReference)
                {
                    (Worksheet as XLWorksheet).Internals.RowsCollection[this.RowNumber()].Style = value;
                }
                else
                {
                    style = new XLStyle(this, value);


                    Int32 minColumn = 1;
                    Int32 maxColumn = 0;
                    var row = this.RowNumber();
                    if ((Worksheet as XLWorksheet).Internals.CellsCollection.Values.Any(c => c.Address.RowNumber == row))
                    {
                        minColumn = (Worksheet as XLWorksheet).Internals.CellsCollection.Values
                            .Where(c => c.Address.RowNumber == row)
                            .Min(c => c.Address.ColumnNumber);
                        maxColumn = (Worksheet as XLWorksheet).Internals.CellsCollection.Values
                            .Where(c => c.Address.RowNumber == row)
                            .Max(c => c.Address.ColumnNumber);
                    }

                    if ((Worksheet as XLWorksheet).Internals.ColumnsCollection.Count > 0)
                    {
                        Int32 minInCollection = (Worksheet as XLWorksheet).Internals.ColumnsCollection.Keys.Min();
                        Int32 maxInCollection = (Worksheet as XLWorksheet).Internals.ColumnsCollection.Keys.Max();
                        if (minInCollection < minColumn) minColumn = minInCollection;
                        if (maxInCollection > maxColumn) maxColumn = maxInCollection;
                    }
                    
                    for (Int32 co = minColumn; co <= maxColumn; co++)
                    {
                        Worksheet.Cell(row, co).Style = value;
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

                var row = this.RowNumber();
                Int32 minColumn = 1;
                Int32 maxColumn = 0;
                if ((Worksheet as XLWorksheet).Internals.CellsCollection.Values.Any(c => c.Address.RowNumber == row))
                    maxColumn = (Worksheet as XLWorksheet).Internals.CellsCollection.Values.Where(c => c.Address.RowNumber == row).Max(c => c.Address.ColumnNumber);

                if ((Worksheet as XLWorksheet).Internals.ColumnsCollection.Count > 0)
                {
                    Int32 maxInCollection = (Worksheet as XLWorksheet).Internals.ColumnsCollection.Keys.Max();
                    if (maxInCollection > maxColumn) maxColumn = maxInCollection;
                }

                for (var co = minColumn; co <= maxColumn; co++)
                {
                    yield return Worksheet.Cell(row, co).Style;
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
                    return (Worksheet as XLWorksheet).Internals.RowsCollection[this.RowNumber()].InnerStyle;
                else
                    return new XLStyle(new XLStylizedContainer(this.style, this), style);
            }
            set
            {
                if (IsReference)
                {
                    (Worksheet as XLWorksheet).Internals.RowsCollection[this.RowNumber()].InnerStyle = value;
                }
                else
                {
                    style = new XLStyle(this, value);
                }
            }
        }

        public override IXLRange AsRange()
        {
            return Range(1, 1, 1, XLWorksheet.MaxNumberOfColumns);
        }

        #endregion

        private Boolean collapsed;
        public Boolean Collapsed
        {
            get
            {
                if (IsReference)
                {
                    return (Worksheet as XLWorksheet).Internals.RowsCollection[this.RowNumber()].Collapsed;
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
                    (Worksheet as XLWorksheet).Internals.RowsCollection[this.RowNumber()].Collapsed = value;
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
                    return (Worksheet as XLWorksheet).Internals.RowsCollection[this.RowNumber()].OutlineLevel;
                }
                else
                {
                    return outlineLevel;
                }
            }
            set
            {
                if (value < 1 || value > 8)
                    throw new ArgumentOutOfRangeException("Outline level must be between 1 and 8.");

                if (IsReference)
                {
                    (Worksheet as XLWorksheet).Internals.RowsCollection[this.RowNumber()].OutlineLevel = value;
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

        public IXLRow Sort()
        {
            this.RangeUsed().Sort(XLSortOrientation.LeftToRight);
            return this;
        }
        public IXLRow Sort(XLSortOrder sortOrder)
        {
            this.RangeUsed().Sort(XLSortOrientation.LeftToRight, sortOrder);
            return this;
        }
        public IXLRow Sort(Boolean matchCase)
        {
            this.AsRange().Sort(XLSortOrientation.LeftToRight, matchCase);
            return this;
        }
        public IXLRow Sort(XLSortOrder sortOrder, Boolean matchCase)
        {
            this.AsRange().Sort(XLSortOrientation.LeftToRight, sortOrder, matchCase);
            return this;
        }

        private void CopyToCell(IXLRangeRow rngRow, IXLCell cell)
        {
            Int32 cellCount = rngRow.CellCount();
            Int32 roStart = cell.Address.RowNumber;
            Int32 coStart = cell.Address.ColumnNumber;
            for (Int32 co = coStart; co <= cellCount + coStart - 1; co++)
            {
                (cell.Worksheet.Cell(roStart, co) as XLCell).CopyFrom(rngRow.Cell(co - coStart + 1));
            }
        }

        public new IXLRangeRow CopyTo(IXLCell target)
        {
            var rngUsed = RangeUsed().Row(1);
            CopyToCell(rngUsed, target);
            
            Int32 lastColumnNumber =  target.Address.ColumnNumber + rngUsed.CellCount() - 1;
            if (lastColumnNumber > XLWorksheet.MaxNumberOfColumns) lastColumnNumber = XLWorksheet.MaxNumberOfColumns;

            return target.Worksheet.Range(
                target.Address.RowNumber,
                target.Address.ColumnNumber,
                target.Address.RowNumber,
                lastColumnNumber)
                .Row(1);
        }
        public new IXLRangeRow CopyTo(IXLRangeBase target)
        {
            var thisRangeUsed = RangeUsed();
            Int32 thisColumnCount = thisRangeUsed.ColumnCount();
            var targetRangeUsed = target.AsRange().RangeUsed();
            Int32 targetColumnCount = targetRangeUsed.ColumnCount();
            Int32 maxColumn = thisColumnCount > targetColumnCount ? thisColumnCount : targetColumnCount;

            CopyToCell(this.Range(1, 1, 1, maxColumn).Row(1), target.FirstCell());

            Int32 lastColumnNumber = target.RangeAddress.LastAddress.ColumnNumber + maxColumn - 1;
            if (lastColumnNumber > XLWorksheet.MaxNumberOfColumns) lastColumnNumber = XLWorksheet.MaxNumberOfColumns;

            return (target as XLRangeBase).Worksheet.Range(
                target.RangeAddress.FirstAddress.RowNumber,
                target.RangeAddress.LastAddress.ColumnNumber,
                target.RangeAddress.FirstAddress.RowNumber,
                lastColumnNumber)
                .Row(1);
        }
        public IXLRow CopyTo(IXLRow row)
        {
            var thisRangeUsed = RangeUsed();
            Int32 thisColumnCount = thisRangeUsed.ColumnCount();
            //var targetRangeUsed = column target.AsRange().RangeUsed();
            Int32 targetColumnCount = row.LastCellUsed(true).Address.ColumnNumber;
            Int32 maxColumn = thisColumnCount > targetColumnCount ? thisColumnCount : targetColumnCount;

            CopyToCell(this.Row(1, maxColumn), row.FirstCell());
            var newRow = row as XLRow;
            newRow.height = height;
            newRow.style = new XLStyle(newRow, Style);
            return newRow;
        }

        public IXLRangeRow Row(Int32 start, Int32 end)
        {
            return Range(1, start, 1, end).Row(1);
        }
        public IXLRangeRows Rows(String rows)
        {
            var retVal = new XLRangeRows();
            var rowPairs = rows.Split(',');
            foreach (var pair in rowPairs)
            {
                this.AsRange().Rows(pair.Trim()).ForEach(r => retVal.Add(r));
            }
            return retVal;
        }

        public IXLRow AddHorizontalPageBreak()
        {
            Worksheet.PageSetup.AddHorizontalPageBreak(this.RowNumber());
            return this;
        }
    }
}
