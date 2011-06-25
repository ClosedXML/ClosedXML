using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLRow : XLRangeBase, IXLRow
    {
        #region Static
        private static double DegreeToRadian(double angle)
        {
            return Math.PI*angle/180.0;
        }
        #endregion
        #region Private fields
        private Boolean m_collapsed;
        private Int32 m_outlineLevel;
        private Double m_height;
        private Boolean m_isHidden;
        private IXLStyle style;
        #endregion
        #region Constructor
        public XLRow(Int32 row, XLRowParameters xlRowParameters)
                : base(new XLRangeAddress(new XLAddress(xlRowParameters.Worksheet, row, 1, false, false),
                                          new XLAddress(xlRowParameters.Worksheet, row, ExcelHelper.MaxColumnNumber, false, false)))
        {
            SetRowNumber(row);

            IsReference = xlRowParameters.IsReference;
            if (IsReference)
            {
                //SMELL: Leak may occur
                (Worksheet).RangeShiftedRows += Worksheet_RangeShiftedRows;
            }
            else
            {
                style = new XLStyle(this, xlRowParameters.DefaultStyle);
                m_height = xlRowParameters.Worksheet.RowHeight;
            }
        }

        public XLRow(XLRow row)
                : base(new XLRangeAddress(new XLAddress(row.Worksheet, row.RowNumber(), 1, false, false),
                                          new XLAddress(row.Worksheet, row.RowNumber(), ExcelHelper.MaxColumnNumber, false, false)))
        {
            m_height = row.m_height;
            IsReference = row.IsReference;
            m_collapsed = row.m_collapsed;
            m_isHidden = row.m_isHidden;
            m_outlineLevel = row.m_outlineLevel;
            style = new XLStyle(this, row.Style);
        }
        #endregion
        private void Worksheet_RangeShiftedRows(XLRange range, int rowsShifted)
        {
            if (range.RangeAddress.FirstAddress.RowNumber <= RowNumber())
            {
                SetRowNumber(RowNumber() + rowsShifted);
            }
        }

        private void RowsCollection_RowShifted(int startingRow, int rowsShifted)
        {
            if (startingRow <= RowNumber())
            {
                SetRowNumber(RowNumber() + rowsShifted);
            }
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
                RangeAddress.LastAddress = new XLAddress(Worksheet,
                                                         row,
                                                         ExcelHelper.MaxColumnNumber,
                                                         RangeAddress.LastAddress.FixedRow,
                                                         RangeAddress.LastAddress.FixedColumn);
            }
        }

        public Boolean IsReference { get; private set; }
        #region IXLRow Members
        public Double Height
        {
            get
            {
                if (IsReference)
                {
                    return (Worksheet).Internals.RowsCollection[RowNumber()].Height;
                }
                return m_height;
            }
            set
            {
                if (IsReference)
                {
                    (Worksheet).Internals.RowsCollection[RowNumber()].Height = value;
                }
                else
                {
                    m_height = value;
                }
            }
        }

        public void Delete()
        {
            var rowNumber = RowNumber();
            AsRange().Delete(XLShiftDeletedCells.ShiftCellsUp);
            (Worksheet).Internals.RowsCollection.Remove(rowNumber);
            var rowsToMove = new List<Int32>();
            rowsToMove.AddRange((Worksheet).Internals.RowsCollection.Where(c => c.Key > rowNumber).Select(c => c.Key));
            foreach (var row in rowsToMove.OrderBy(r => r))
            {
                (Worksheet).Internals.RowsCollection.Add(row - 1, (Worksheet).Internals.RowsCollection[row]);
                (Worksheet).Internals.RowsCollection.Remove(row);
            }
        }

        public new IXLRows InsertRowsBelow(Int32 numberOfRows)
        {
            var rowNum = RowNumber();
            (Worksheet).Internals.RowsCollection.ShiftRowsDown(rowNum + 1, numberOfRows);
            XLRange range = (XLRange) Worksheet.Row(rowNum).AsRange();
            range.InsertRowsBelow(true, numberOfRows);
            return Worksheet.Rows(rowNum + 1, rowNum + numberOfRows);
        }

        public new IXLRows InsertRowsAbove(Int32 numberOfRows)
        {
            var rowNum = RowNumber();
            (Worksheet).Internals.RowsCollection.ShiftRowsDown(rowNum, numberOfRows);
            // We can't use this.AsRange() because we've shifted the rows
            // and we want to use the old rowNum.
            XLRange range = (XLRange) Worksheet.Row(rowNum).AsRange();
            range.InsertRowsAbove(true, numberOfRows);
            return Worksheet.Rows(rowNum, rowNum + numberOfRows - 1);
        }

        public new void Clear()
        {
            var range = AsRange();
            range.Clear();
            Style = Worksheet.Style;
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

        public override XLRange Range(String rangeAddressStr)
        {
            String rangeAddressToUse;
            if (rangeAddressStr.Contains(':') || rangeAddressStr.Contains('-'))
            {
                if (rangeAddressStr.Contains('-'))
                {
                    rangeAddressStr = rangeAddressStr.Replace('-', ':');
                }

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
            return Cells(ExcelHelper.GetColumnNumberFromLetter(firstColumn) + ":"
                         + ExcelHelper.GetColumnNumberFromLetter(lastColumn));
        }
        public IXLRow AdjustToContents()
        {
            return AdjustToContents(1);
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
            foreach (var cell in Row(startColumn, endColumn).CellsUsed())
            {
                var c = (XLCell) cell;
                var cellAsRange = c.AsRange();
                Boolean isMerged = Worksheet.Internals.MergedRanges.Any(m => cellAsRange.Intersects(m));
                if (!isMerged)
                {
                    Double thisHeight;
                    Int32 textRotation = c.Style.Alignment.TextRotation;
                    if (c.HasRichText || textRotation != 0 || c.InnerText.Contains(Environment.NewLine))
                    {
                        var kpList = new List<KeyValuePair<IXLFontBase, string>>();
                        if (c.HasRichText)
                        {
                            foreach (var rt in c.RichText)
                            {
                                String formattedString = rt.Text;
                                var arr = formattedString.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
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
                            var arr = formattedString.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
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
                        {
                            thisHeight = maxHeightCol * lineCount;
                        }
                        else
                        {
                            if (textRotation == 255)
                            {
                                thisHeight = maxLongCol * maxHeightCol;
                            }
                            else
                            {
                                Double rotation;
                                if (textRotation == 90 || textRotation == 180 || textRotation == 255)
                                {
                                    rotation = 90;
                                }
                                else
                                {
                                    rotation = textRotation % 90;
                                }

                                thisHeight = (rotation / 90.0) * maxHeightCol * maxLongCol * 0.5;
                            }
                        }
                    }
                    else
                    {
                        thisHeight = c.Style.Font.GetHeight();
                    }

                    if (thisHeight >= maxHeight)
                    {
                        rowMaxHeight = maxHeight;
                        break;
                    }
                    if (thisHeight > rowMaxHeight)
                    {
                        rowMaxHeight = thisHeight;
                    }
                }
            }

            if (rowMaxHeight == 0)
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
            get
            {
                return IsReference ? (Worksheet).Internals.RowsCollection[RowNumber()].IsHidden : m_isHidden;
            }
            set
            {
                if (IsReference)
                {
                    (Worksheet).Internals.RowsCollection[RowNumber()].IsHidden = value;
                }
                else
                {
                    m_isHidden = value;
                }
            }
        }
        #endregion
        #region IXLStylized Members
        internal void SetStyleNoColumns(IXLStyle value)
        {
            if (IsReference)
            {
                (Worksheet).Internals.RowsCollection[RowNumber()].SetStyleNoColumns(value);
            }
            else
            {
                style = new XLStyle(this, value);

                var row = RowNumber();
                foreach (var c in (Worksheet).Internals.CellsCollection.Values.Where(c => c.Address.RowNumber == row))
                {
                    c.Style = value;
                }
            }
        }

        public override IXLStyle Style
        {
            get
            {
                if (IsReference)
                {
                    return (Worksheet).Internals.RowsCollection[RowNumber()].Style;
                }
                else
                {
                    return style;
                }
            }
            set
            {
                if (IsReference)
                {
                    (Worksheet).Internals.RowsCollection[RowNumber()].Style = value;
                }
                else
                {
                    style = new XLStyle(this, value);

                    Int32 minColumn = 1;
                    Int32 maxColumn = 0;
                    var row = RowNumber();
                    if ((Worksheet).Internals.CellsCollection.Values.Any(c => c.Address.RowNumber == row))
                    {
                        minColumn = (Worksheet).Internals.CellsCollection.Values
                                .Where(c => c.Address.RowNumber == row)
                                .Min(c => c.Address.ColumnNumber);
                        maxColumn = (Worksheet).Internals.CellsCollection.Values
                                .Where(c => c.Address.RowNumber == row)
                                .Max(c => c.Address.ColumnNumber);
                    }

                    if ((Worksheet).Internals.ColumnsCollection.Count > 0)
                    {
                        Int32 minInCollection = (Worksheet).Internals.ColumnsCollection.Keys.Min();
                        Int32 maxInCollection = (Worksheet).Internals.ColumnsCollection.Keys.Max();
                        if (minInCollection < minColumn)
                        {
                            minColumn = minInCollection;
                        }
                        if (maxInCollection > maxColumn)
                        {
                            maxColumn = maxInCollection;
                        }
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

                var row = RowNumber();
                Int32 minColumn = 1;
                Int32 maxColumn = 0;
                if ((Worksheet).Internals.CellsCollection.Values.Any(c => c.Address.RowNumber == row))
                {
                    maxColumn =
                            (Worksheet).Internals.CellsCollection.Values.Where(c => c.Address.RowNumber == row).Max(
                                    c => c.Address.ColumnNumber);
                }

                if ((Worksheet).Internals.ColumnsCollection.Count > 0)
                {
                    Int32 maxInCollection = (Worksheet).Internals.ColumnsCollection.Keys.Max();
                    if (maxInCollection > maxColumn)
                    {
                        maxColumn = maxInCollection;
                    }
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
                {
                    return (Worksheet).Internals.RowsCollection[RowNumber()].InnerStyle;
                }
                return new XLStyle(new XLStylizedContainer(style, this), style);
            }
            set
            {
                if (IsReference)
                {
                    (Worksheet).Internals.RowsCollection[RowNumber()].InnerStyle = value;
                }
                else
                {
                    style = new XLStyle(this, value);
                }
            }
        }

        public override IXLRange AsRange()
        {
            return Range(1, 1, 1, ExcelHelper.MaxColumnNumber);
        }
        #endregion
        public Boolean Collapsed
        {
            get
            {
                if (IsReference)
                {
                    return (Worksheet).Internals.RowsCollection[RowNumber()].Collapsed;
                }
                else
                {
                    return m_collapsed;
                }
            }
            set
            {
                if (IsReference)
                {
                    (Worksheet).Internals.RowsCollection[RowNumber()].Collapsed = value;
                }
                else
                {
                    m_collapsed = value;
                }
            }
        }

        public Int32 OutlineLevel
        {
            get
            {
                if (IsReference)
                {
                    return (Worksheet).Internals.RowsCollection[RowNumber()].OutlineLevel;
                }
                return m_outlineLevel;
            }
            set
            {
                if (value < 1 || value > 8)
                {
                    throw new ArgumentOutOfRangeException("value", "Outline level must be between 1 and 8.");
                }

                if (IsReference)
                {
                    Worksheet.Internals.RowsCollection[RowNumber()].OutlineLevel = value;
                }
                else
                {
                    Worksheet.IncrementColumnOutline(value);
                    Worksheet.DecrementColumnOutline(m_outlineLevel);
                    m_outlineLevel = value;
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
            {
                OutlineLevel += 1;
            }

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
                {
                    OutlineLevel -= 1;
                }
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

        private void CopyToCell(IXLRangeRow rngRow, XLCell cell)
        {
            Int32 cellCount = rngRow.CellCount();
            Int32 roStart = cell.Address.RowNumber;
            Int32 coStart = cell.Address.ColumnNumber;
            for (Int32 co = coStart; co <= cellCount + coStart - 1; co++)
            {
                (cell.Worksheet.Cell(roStart, co)).CopyFrom((XLCell) rngRow.Cell(co - coStart + 1));
            }
        }

        IXLRangeRow IXLRow.CopyTo(IXLCell target)
        {
            var rngUsed = RangeUsed(true).Row(1);
            CopyToCell(rngUsed, (XLCell) target);

            Int32 lastColumnNumber = target.Address.ColumnNumber + rngUsed.CellCount() - 1;
            if (lastColumnNumber > ExcelHelper.MaxColumnNumber)
            {
                lastColumnNumber = ExcelHelper.MaxColumnNumber;
            }

            return target.Worksheet.Range(
                    target.Address.RowNumber,
                    target.Address.ColumnNumber,
                    target.Address.RowNumber,
                    lastColumnNumber)
                    .Row(1);
        }
        public override void CopyTo(IXLCell target)
        {
            ((IXLRow)this).CopyTo(target);
        }
        IXLRangeRow IXLRow.CopyTo(IXLRangeBase target)
        {
            var thisRangeUsed = RangeUsed(true);
            Int32 thisColumnCount = thisRangeUsed.ColumnCount();
            var targetRangeUsed = target.AsRange().RangeUsed();
            Int32 targetColumnCount = targetRangeUsed.ColumnCount();
            Int32 maxColumn = thisColumnCount > targetColumnCount ? thisColumnCount : targetColumnCount;

            CopyToCell(Range(1, 1, 1, maxColumn).Row(1), (XLCell)target.FirstCell());

            Int32 lastColumnNumber = target.RangeAddress.LastAddress.ColumnNumber + maxColumn - 1;
            if (lastColumnNumber > ExcelHelper.MaxColumnNumber)
            {
                lastColumnNumber = ExcelHelper.MaxColumnNumber;
            }

            return target.Worksheet.Range(target.RangeAddress.FirstAddress.RowNumber,
                                          target.RangeAddress.LastAddress.ColumnNumber,
                                          target.RangeAddress.FirstAddress.RowNumber,
                                          lastColumnNumber)
                    .Row(1);
        }
        public override void CopyTo(IXLRangeBase target)
        {
            ((IXLRow) this).CopyTo(target);
        }

        public IXLRow CopyTo(IXLRow row)
        {
            row.Clear();
            var originalRange = RangeUsed(true);
            if (!ReferenceEquals(originalRange, null))
            {
                int columnNumber = originalRange.ColumnCount();
                var destRange = row.Worksheet.Range(row.RowNumber(), ExcelHelper.MinColumnNumber, row.RowNumber(), columnNumber);
                originalRange.CopyTo(destRange);
                //Old
                //CopyToCell(Row(1, columnNumber), (XLCell) row.FirstCell());
            }
            var newRow = (XLRow) row;
            newRow.m_height = m_height;
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
                AsRange().Rows(pair.Trim()).ForEach(retVal.Add);
            }
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
    }
}