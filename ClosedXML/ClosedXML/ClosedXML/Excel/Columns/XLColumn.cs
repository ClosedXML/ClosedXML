using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLColumn : XLRangeBase, IXLColumn
    {
        public XLColumn(Int32 column, XLColumnParameters xlColumnParameters)
                : base(
                        new XLRangeAddress(new XLAddress(xlColumnParameters.Worksheet, 1, column, false, false),
                                           new XLAddress(xlColumnParameters.Worksheet, XLWorksheet.MaxNumberOfRows, column, false, false)))
        {
            SetColumnNumber(column);

            IsReference = xlColumnParameters.IsReference;
            if (IsReference)
            {
                (Worksheet).RangeShiftedColumns += Worksheet_RangeShiftedColumns;
            }
            else
            {
                style = new XLStyle(this, xlColumnParameters.DefaultStyle);
                width = xlColumnParameters.Worksheet.ColumnWidth;
            }
        }

        public XLColumn(XLColumn column)
                : base(
                        new XLRangeAddress(new XLAddress(column.Worksheet, 1, column.ColumnNumber(), false, false),
                                           new XLAddress(column.Worksheet, XLWorksheet.MaxNumberOfRows, column.ColumnNumber(), false, false)))
        {
            width = column.width;
            IsReference = column.IsReference;
            collapsed = column.collapsed;
            isHidden = column.isHidden;
            outlineLevel = column.outlineLevel;
            style = new XLStyle(this, column.Style);
        }

        private void Worksheet_RangeShiftedColumns(XLRange range, int columnsShifted)
        {
            if (range.RangeAddress.FirstAddress.ColumnNumber <= ColumnNumber())
            {
                SetColumnNumber(ColumnNumber() + columnsShifted);
            }
        }

        private void SetColumnNumber(Int32 column)
        {
            if (column <= 0)
            {
                RangeAddress.IsInvalid = false;
            }
            else
            {
                RangeAddress.FirstAddress = new XLAddress(Worksheet,
                                                          1,
                                                          column,
                                                          RangeAddress.FirstAddress.FixedRow,
                                                          RangeAddress.FirstAddress.FixedColumn);
                RangeAddress.LastAddress = new XLAddress(Worksheet,
                                                         XLWorksheet.MaxNumberOfRows,
                                                         column,
                                                         RangeAddress.LastAddress.FixedRow,
                                                         RangeAddress.LastAddress.FixedColumn);
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
                    return (Worksheet).Internals.ColumnsCollection[ColumnNumber()].Width;
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
                    (Worksheet).Internals.ColumnsCollection[ColumnNumber()].Width = value;
                }
                else
                {
                    width = value;
                }
            }
        }

        public void Delete()
        {
            var columnNumber = ColumnNumber();
            AsRange().Delete(XLShiftDeletedCells.ShiftCellsLeft);
            (Worksheet).Internals.ColumnsCollection.Remove(columnNumber);
            List<Int32> columnsToMove = new List<Int32>();
            columnsToMove.AddRange((Worksheet).Internals.ColumnsCollection.Where(c => c.Key > columnNumber).Select(c => c.Key));
            foreach (var column in columnsToMove.OrderBy(c => c))
            {
                (Worksheet).Internals.ColumnsCollection.Add(column - 1, (Worksheet).Internals.ColumnsCollection[column]);
                (Worksheet).Internals.ColumnsCollection.Remove(column);
            }
        }

        public new void Clear()
        {
            var range = AsRange();
            range.Clear();
            Style = Worksheet.Style;
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
                {
                    return (Worksheet).Internals.ColumnsCollection[ColumnNumber()].Style;
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
                    (Worksheet).Internals.ColumnsCollection[ColumnNumber()].Style = value;
                }
                else
                {
                    style = new XLStyle(this, value);

                    Int32 minRow = 1;
                    Int32 maxRow = 0;
                    var column = ColumnNumber();
                    if ((Worksheet).Internals.CellsCollection.Values.Any(c => c.Address.ColumnNumber == column))
                    {
                        minRow = (Worksheet).Internals.CellsCollection.Values
                                .Where(c => c.Address.ColumnNumber == column)
                                .Min(c => c.Address.RowNumber);
                        maxRow = (Worksheet).Internals.CellsCollection.Values
                                .Where(c => c.Address.ColumnNumber == column)
                                .Max(c => c.Address.RowNumber);
                    }

                    if ((Worksheet).Internals.RowsCollection.Count > 0)
                    {
                        Int32 minInCollection = (Worksheet).Internals.RowsCollection.Keys.Min();
                        Int32 maxInCollection = (Worksheet).Internals.RowsCollection.Keys.Max();
                        if (minInCollection < minRow)
                        {
                            minRow = minInCollection;
                        }
                        if (maxInCollection > maxRow)
                        {
                            maxRow = maxInCollection;
                        }
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

                var column = ColumnNumber();
                Int32 minRow = 1;
                Int32 maxRow = 0;
                if ((Worksheet).Internals.CellsCollection.Values.Any(c => c.Address.ColumnNumber == column))
                {
                    maxRow = (Worksheet).Internals.CellsCollection.Values.Where(c => c.Address.ColumnNumber == column).Max(c => c.Address.RowNumber);
                }

                if ((Worksheet).Internals.RowsCollection.Count > 0)
                {
                    Int32 maxInCollection = (Worksheet).Internals.RowsCollection.Keys.Max();
                    if (maxInCollection > maxRow)
                    {
                        maxRow = maxInCollection;
                    }
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
                {
                    return (Worksheet).Internals.ColumnsCollection[ColumnNumber()].InnerStyle;
                }
                else
                {
                    return new XLStyle(new XLStylizedContainer(style, this), style);
                }
            }
            set
            {
                if (IsReference)
                {
                    (Worksheet).Internals.ColumnsCollection[ColumnNumber()].InnerStyle = value;
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
            var columnNum = ColumnNumber();
            (Worksheet).Internals.ColumnsCollection.ShiftColumnsRight(columnNum + 1, numberOfColumns);
            XLRange range = (XLRange) Worksheet.Column(columnNum).AsRange();
            range.InsertColumnsAfter(true, numberOfColumns);
            return Worksheet.Columns(columnNum + 1, columnNum + numberOfColumns);
        }
        public new IXLColumns InsertColumnsBefore(Int32 numberOfColumns)
        {
            var columnNum = ColumnNumber();
            (Worksheet).Internals.ColumnsCollection.ShiftColumnsRight(columnNum, numberOfColumns);
            // We can't use this.AsRange() because we've shifted the columns
            // and we want to use the old columnNum.
            XLRange range = (XLRange) Worksheet.Column(columnNum).AsRange();
            range.InsertColumnsBefore(true, numberOfColumns);
            return Worksheet.Columns(columnNum, columnNum + numberOfColumns - 1);
        }

        public override IXLRange AsRange()
        {
            return Range(1, 1, XLWorksheet.MaxNumberOfRows, 1);
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
            return AdjustToContents(startRow, endRow, 0, Double.MaxValue);
        }

        private double DegreeToRadian(double angle)
        {
            return Math.PI * angle / 180.0;
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
            foreach (var cell in Column(startRow, endRow).CellsUsed())
            {
                var c = cell as XLCell;
                Boolean isMerged = CellIsMerged(c);

                if (!isMerged)
                {
                    Double thisWidthMax = 0;
                    Int32 textRotation = c.Style.Alignment.TextRotation;
                    if (c.HasRichText || textRotation != 0 || c.InnerText.Contains(Environment.NewLine))
                    {

                        List<KeyValuePair<IXLFontBase, String>> kpList = new List<KeyValuePair<IXLFontBase, string>>();

                        #region if (c.HasRichText)

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
                        #endregion

                        #region foreach (var kp in kpList)


                        Double runningWidth = 0;
                        Boolean rotated = false;
                        Double maxLineWidth = 0;
                        Int32 lineCount = 1;
                        foreach (var kp in kpList)
                        {
                            IXLFontBase f = kp.Key;
                            String formattedString = kp.Value;

                            Int32 newLinePosition = formattedString.IndexOf(Environment.NewLine);
                            if (textRotation == 0)
                            {
                                #region if (newLinePosition >= 0)

                                if (newLinePosition >= 0)
                                {
                                    if (newLinePosition > 0)
                                        runningWidth += f.GetWidth(formattedString.Substring(0, newLinePosition));

                                    if (runningWidth > thisWidthMax)
                                        thisWidthMax = runningWidth;

                                    if (newLinePosition < formattedString.Length - 2)
                                        runningWidth = f.GetWidth(formattedString.Substring(newLinePosition + 2));
                                    else
                                        runningWidth = 0;
                                }
                                else
                                {
                                    runningWidth += f.GetWidth(formattedString);
                                }
                                #endregion
                            }
                            else
                            {
                                #region if (textRotation == 255)
                                if (textRotation == 255)
                                {
                                    if (runningWidth == 0)
                                        runningWidth = f.GetWidth("X");

                                    if (newLinePosition >= 0)
                                        runningWidth += f.GetWidth("X");
                                }
                                else
                                {
                                    rotated = true;
                                    Double vWidth = f.GetWidth("X");
                                    if (vWidth > maxLineWidth)
                                        maxLineWidth = vWidth;

                                    if (newLinePosition >= 0)
                                    {
                                        lineCount++;

                                        if (newLinePosition > 0)
                                            runningWidth += f.GetWidth(formattedString.Substring(0, newLinePosition));

                                        if (runningWidth > thisWidthMax)
                                            thisWidthMax = runningWidth;

                                        if (newLinePosition < formattedString.Length - 2)
                                            runningWidth = f.GetWidth(formattedString.Substring(newLinePosition + 2));
                                        else
                                            runningWidth = 0;

                                    }
                                    else
                                    {
                                        runningWidth += f.GetWidth(formattedString);
                                    }
                                }
                                #endregion
                            }
                        }
                        #endregion
                        if (runningWidth > thisWidthMax)
                            thisWidthMax = runningWidth;

                        #region if (rotated)
                        if (rotated)
                        {
                            Int32 rotation;
                            if (textRotation == 90 || textRotation == 180 || textRotation == 255)
                                rotation = 90;
                            else
                                rotation = textRotation % 90;

                            Double r = DegreeToRadian(rotation);
  
                            thisWidthMax = (thisWidthMax  * Math.Cos(r)) + (maxLineWidth * lineCount) ;
                        }
                        #endregion
                    }
                    else
                    {
                        thisWidthMax = c.Style.Font.GetWidth(c.GetFormattedString());
                    }
                    if (thisWidthMax >= maxWidth)
                    {
                        colMaxWidth = maxWidth;
                        break;
                    }
                    else if (thisWidthMax > colMaxWidth)
                    {
                        colMaxWidth = thisWidthMax + 1;
                    }
                }
            }

            if (colMaxWidth == 0)
                colMaxWidth = Worksheet.ColumnWidth;

            Width = colMaxWidth;

            return this;
        }

        private Boolean CellIsMerged(IXLCell c)
        {
            Boolean isMerged = false;
            var cellAsRange = c.AsRange();
            foreach (var m in Worksheet.Internals.MergedRanges)
            {
                if (cellAsRange.Intersects(m))
                {
                    isMerged = true;
                    break;
                }
            }
            return isMerged;
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
                    return (Worksheet).Internals.ColumnsCollection[ColumnNumber()].IsHidden;
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
                    (Worksheet).Internals.ColumnsCollection[ColumnNumber()].IsHidden = value;
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
                return IsReference ? (Worksheet).Internals.ColumnsCollection[ColumnNumber()].Collapsed : collapsed;
            }
            set
            {
                if (IsReference)
                {
                    (Worksheet).Internals.ColumnsCollection[ColumnNumber()].Collapsed = value;
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
            get {
                return IsReference ? (Worksheet).Internals.ColumnsCollection[ColumnNumber()].OutlineLevel : outlineLevel;
            }
            set
            {
                if (value < 0 || value > 8)
                {
                    throw new ArgumentOutOfRangeException("value", "Outline level must be between 0 and 8.");
                }

                if (IsReference)
                {
                    (Worksheet).Internals.ColumnsCollection[ColumnNumber()].OutlineLevel = value;
                }
                else
                {
                    (Worksheet).IncrementColumnOutline(value);
                    (Worksheet).DecrementColumnOutline(outlineLevel);
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
            {
                OutlineLevel += 1;
            }

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

        public IXLColumn Sort()
        {
            RangeUsed().Sort();
            return this;
        }
        public IXLColumn Sort(XLSortOrder sortOrder)
        {
            RangeUsed().Sort(sortOrder);
            return this;
        }
        public IXLColumn Sort(Boolean matchCase)
        {
            AsRange().Sort(matchCase);
            return this;
        }
        public IXLColumn Sort(XLSortOrder sortOrder, Boolean matchCase)
        {
            AsRange().Sort(sortOrder, matchCase);
            return this;
        }

        private void CopyToCell(IXLRangeColumn rngColumn, XLCell cell)
        {
            Int32 cellCount = rngColumn.CellCount();
            Int32 roStart = cell.Address.RowNumber;
            Int32 coStart = cell.Address.ColumnNumber;
            for (Int32 ro = roStart; ro <= cellCount + roStart - 1; ro++)
            {
                cell.Worksheet.Cell(ro, coStart).CopyFrom((XLCell) rngColumn.Cell(ro - roStart + 1));
            }
        }

        public new IXLRangeColumn CopyTo(IXLCell target)
        {
            var rngUsed = RangeUsed().Column(1);
            CopyToCell(rngUsed, (XLCell)target);

            Int32 lastRowNumber = target.Address.RowNumber + rngUsed.CellCount() - 1;
            if (lastRowNumber > XLWorksheet.MaxNumberOfRows)
            {
                lastRowNumber = XLWorksheet.MaxNumberOfRows;
            }

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

            CopyToCell(Range(1, 1, maxRow, 1).Column(1), (XLCell)target.FirstCell());

            Int32 lastRowNumber = target.RangeAddress.FirstAddress.RowNumber + maxRow - 1;
            if (lastRowNumber > XLWorksheet.MaxNumberOfRows)
            {
                lastRowNumber = XLWorksheet.MaxNumberOfRows;
            }

            return (target as XLRangeBase).Worksheet.Range(
                    target.RangeAddress.FirstAddress.RowNumber,
                    target.RangeAddress.LastAddress.ColumnNumber,
                    lastRowNumber,
                    target.RangeAddress.LastAddress.ColumnNumber)
                    .Column(1);
        }
        public IXLColumn CopyTo(IXLColumn column)
        {
            var thisRangeUsed = RangeUsed();
            Int32 thisRowCount = thisRangeUsed.RowCount();
            //var targetRangeUsed = column target.AsRange().RangeUsed();
            Int32 targetRowCount = column.LastCellUsed(true).Address.RowNumber;
            Int32 maxRow = thisRowCount > targetRowCount ? thisRowCount : targetRowCount;

            CopyToCell(Column(1, maxRow), (XLCell) column.FirstCell());
            var newColumn = (XLColumn) column;
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
                AsRange().Columns(pair.Trim()).ForEach(retVal.Add);
            }
            return retVal;
        }

        /// <summary>
        /// 	Adds a vertical page break after this column.
        /// </summary>
        public IXLColumn AddVerticalPageBreak()
        {
            Worksheet.PageSetup.AddVerticalPageBreak(ColumnNumber());
            return this;
        }

        public IXLColumn SetDataType(XLCellValues dataType)
        {
            DataType = dataType;
            return this;
        }
    }
}