using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLColumn : XLRangeBase, IXLColumn
    {
        #region Private fields

        private bool _collapsed;
        private bool _isHidden;
        private int _outlineLevel;

        private Double _width;

        #endregion Private fields

        #region Constructor

        public XLColumn(Int32 column, XLColumnParameters xlColumnParameters)
            : base(
                new XLRangeAddress(new XLAddress(xlColumnParameters.Worksheet, 1, column, false, false),
                                   new XLAddress(xlColumnParameters.Worksheet, XLHelper.MaxRowNumber, column, false,
                                                 false)),
                xlColumnParameters.IsReference ? xlColumnParameters.Worksheet.Internals.ColumnsCollection[column].StyleValue
                                               : (xlColumnParameters.DefaultStyle as XLStyle).Value)
        {
            SetColumnNumber(column);

            IsReference = xlColumnParameters.IsReference;
            if (IsReference)
                SubscribeToShiftedColumns((range, columnsShifted) => this.WorksheetRangeShiftedColumns(range, columnsShifted));
            else
            {
                _width = xlColumnParameters.Worksheet.ColumnWidth;
            }
        }

        public XLColumn(XLColumn column)
            : base(
                new XLRangeAddress(new XLAddress(column.Worksheet, 1, column.ColumnNumber(), false, false),
                                   new XLAddress(column.Worksheet, XLHelper.MaxRowNumber, column.ColumnNumber(),
                                                 false, false)),
                column.StyleValue)
        {
            _width = column._width;
            IsReference = column.IsReference;
            if (IsReference)
                SubscribeToShiftedColumns((range, columnsShifted) => this.WorksheetRangeShiftedColumns(range, columnsShifted));
            _collapsed = column._collapsed;
            _isHidden = column._isHidden;
            _outlineLevel = column._outlineLevel;
        }

        #endregion Constructor

        public Boolean IsReference { get; private set; }

        public override IEnumerable<IXLStyle> Styles
        {
            get
            {
                if (IsReference)
                    yield return Worksheet.Internals.ColumnsCollection[ColumnNumber()].Style;
                else
                    yield return Style;

                int column = ColumnNumber();

                foreach (XLCell cell in Worksheet.Internals.CellsCollection.GetCellsInColumn(column))
                    yield return cell.Style;
            }
        }

        protected override IEnumerable<XLStylizedBase> Children
        {
            get
            {
                int column = ColumnNumber();
                if (IsReference)
                    yield return Worksheet.Internals.ColumnsCollection[column];
                else
                {
                    foreach (XLCell cell in Worksheet.Internals.CellsCollection.GetCellsInColumn(column))
                        yield return cell;
                }
            }
        }

        public Boolean Collapsed
        {
            get { return IsReference ? Worksheet.Internals.ColumnsCollection[ColumnNumber()].Collapsed : _collapsed; }
            set
            {
                if (IsReference)
                    Worksheet.Internals.ColumnsCollection[ColumnNumber()].Collapsed = value;
                else
                    _collapsed = value;
            }
        }

        #region IXLColumn Members

        public Double Width
        {
            get { return IsReference ? Worksheet.Internals.ColumnsCollection[ColumnNumber()].Width : _width; }
            set
            {
                if (IsReference)
                    Worksheet.Internals.ColumnsCollection[ColumnNumber()].Width = value;
                else
                    _width = value;
            }
        }

        public void Delete()
        {
            int columnNumber = ColumnNumber();
            using (var asRange = AsRange())
            {
                asRange.Delete(XLShiftDeletedCells.ShiftCellsLeft);
            }

            Worksheet.Internals.ColumnsCollection.Remove(columnNumber);
            var columnsToMove = new List<Int32>();
            columnsToMove.AddRange(
                Worksheet.Internals.ColumnsCollection.Where(c => c.Key > columnNumber).Select(c => c.Key));
            foreach (int column in columnsToMove.OrderBy(c => c))
            {
                Worksheet.Internals.ColumnsCollection.Add(column - 1, Worksheet.Internals.ColumnsCollection[column]);
                Worksheet.Internals.ColumnsCollection.Remove(column);
            }
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

        public new IXLCells Cells(String cellsInColumn)
        {
            var retVal = new XLCells(false, false);
            var rangePairs = cellsInColumn.Split(',');
            foreach (string pair in rangePairs)
                retVal.Add(Range(pair.Trim()).RangeAddress);
            return retVal;
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
                return Cells(FirstCellUsed().Address.RowNumber, LastCellUsed().Address.RowNumber);
        }

        public IXLCells Cells(Int32 firstRow, Int32 lastRow)
        {
            return Cells(firstRow + ":" + lastRow);
        }

        public new IXLColumns InsertColumnsAfter(Int32 numberOfColumns)
        {
            int columnNum = ColumnNumber();
            Worksheet.Internals.ColumnsCollection.ShiftColumnsRight(columnNum + 1, numberOfColumns);
            using (var column = Worksheet.Column(columnNum))
            {
                using (var asRange = column.AsRange())
                {
                    asRange.InsertColumnsAfterVoid(true, numberOfColumns);
                }
            }

            var newColumns = Worksheet.Columns(columnNum + 1, columnNum + numberOfColumns);
            CopyColumns(newColumns);
            return newColumns;
        }

        public new IXLColumns InsertColumnsBefore(Int32 numberOfColumns)
        {
            int columnNum = ColumnNumber();
            if (columnNum > 1)
            {
                using (var column = Worksheet.Column(columnNum - 1))
                {
                    return column.InsertColumnsAfter(numberOfColumns);
                }
            }

            Worksheet.Internals.ColumnsCollection.ShiftColumnsRight(columnNum, numberOfColumns);

            using (var column = Worksheet.Column(columnNum))
            {
                using (var asRange = column.AsRange())
                {
                    asRange.InsertColumnsBeforeVoid(true, numberOfColumns);
                }
            }

            return Worksheet.Columns(columnNum, columnNum + numberOfColumns - 1);
        }

        private void CopyColumns(IXLColumns newColumns)
        {
            foreach (var newColumn in newColumns)
            {
                var internalColumn = Worksheet.Internals.ColumnsCollection[newColumn.ColumnNumber()];
                internalColumn._width = Width;
                internalColumn.InnerStyle = InnerStyle;
                internalColumn._collapsed = Collapsed;
                internalColumn._isHidden = IsHidden;
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

        public IXLColumn AdjustToContents(Int32 startRow, Int32 endRow, Double minWidth, Double maxWidth)
        {
            var fontCache = new Dictionary<IXLFontBase, Font>();

            Double colMaxWidth = minWidth;

            List<Int32> autoFilterRows = new List<Int32>();
            if (this.Worksheet.AutoFilter != null && this.Worksheet.AutoFilter.Range != null)
                autoFilterRows.Add(this.Worksheet.AutoFilter.Range.FirstRow().RowNumber());

            autoFilterRows.AddRange(Worksheet.Tables.Where(t =>
                    t.AutoFilter != null
                    && t.AutoFilter.Range != null
                    && !autoFilterRows.Contains(t.AutoFilter.Range.FirstRow().RowNumber()))
                .Select(t => t.AutoFilter.Range.FirstRow().RowNumber()));

            XLStyle cellStyle = null;
            foreach (var c in Column(startRow, endRow).CellsUsed().Cast<XLCell>())
            {
                if (c.IsMerged()) continue;
                if (cellStyle == null || cellStyle.Value != c.StyleValue)
                    cellStyle = c.Style as XLStyle;

                Double thisWidthMax = 0;
                Int32 textRotation = cellStyle.Alignment.TextRotation;
                if (c.HasRichText || textRotation != 0 || c.InnerText.Contains(Environment.NewLine))
                {
                    var kpList = new List<KeyValuePair<IXLFontBase, string>>();

                    #region if (c.HasRichText)

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
                            kpList.Add(new KeyValuePair<IXLFontBase, String>(cellStyle.Font, s));
                        }
                    }

                    #endregion if (c.HasRichText)

                    #region foreach (var kp in kpList)

                    Double runningWidth = 0;
                    Boolean rotated = false;
                    Double maxLineWidth = 0;
                    Int32 lineCount = 1;
                    foreach (KeyValuePair<IXLFontBase, string> kp in kpList)
                    {
                        var f = kp.Key;
                        String formattedString = kp.Value;

                        Int32 newLinePosition = formattedString.IndexOf(Environment.NewLine);
                        if (textRotation == 0)
                        {
                            #region if (newLinePosition >= 0)

                            if (newLinePosition >= 0)
                            {
                                if (newLinePosition > 0)
                                    runningWidth += f.GetWidth(formattedString.Substring(0, newLinePosition), fontCache);

                                if (runningWidth > thisWidthMax)
                                    thisWidthMax = runningWidth;

                                runningWidth = newLinePosition < formattedString.Length - 2
                                                   ? f.GetWidth(formattedString.Substring(newLinePosition + 2), fontCache)
                                                   : 0;
                            }
                            else
                                runningWidth += f.GetWidth(formattedString, fontCache);

                            #endregion if (newLinePosition >= 0)
                        }
                        else
                        {
                            #region if (textRotation == 255)

                            if (textRotation == 255)
                            {
                                if (runningWidth <= 0)
                                    runningWidth = f.GetWidth("X", fontCache);

                                if (newLinePosition >= 0)
                                    runningWidth += f.GetWidth("X", fontCache);
                            }
                            else
                            {
                                rotated = true;
                                Double vWidth = f.GetWidth("X", fontCache);
                                if (vWidth > maxLineWidth)
                                    maxLineWidth = vWidth;

                                if (newLinePosition >= 0)
                                {
                                    lineCount++;

                                    if (newLinePosition > 0)
                                        runningWidth += f.GetWidth(formattedString.Substring(0, newLinePosition), fontCache);

                                    if (runningWidth > thisWidthMax)
                                        thisWidthMax = runningWidth;

                                    runningWidth = newLinePosition < formattedString.Length - 2
                                                       ? f.GetWidth(formattedString.Substring(newLinePosition + 2), fontCache)
                                                       : 0;
                                }
                                else
                                    runningWidth += f.GetWidth(formattedString, fontCache);
                            }

                            #endregion if (textRotation == 255)
                        }
                    }

                    #endregion foreach (var kp in kpList)

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

                        thisWidthMax = (thisWidthMax * Math.Cos(r)) + (maxLineWidth * lineCount);
                    }

                    #endregion if (rotated)
                }
                else
                    thisWidthMax = cellStyle.Font.GetWidth(c.GetFormattedString(), fontCache);

                if (autoFilterRows.Contains(c.Address.RowNumber))
                    thisWidthMax += 2.7148; // Allow room for arrow icon in autofilter

                if (thisWidthMax >= maxWidth)
                {
                    colMaxWidth = maxWidth;
                    break;
                }

                if (thisWidthMax > colMaxWidth)
                    colMaxWidth = thisWidthMax + 1;
            }

            if (colMaxWidth <= 0)
                colMaxWidth = Worksheet.ColumnWidth;

            Width = colMaxWidth;

            foreach (IDisposable font in fontCache.Values)
            {
                font.Dispose();
            }
            return this;
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

        public Boolean IsHidden
        {
            get { return IsReference ? Worksheet.Internals.ColumnsCollection[ColumnNumber()].IsHidden : _isHidden; }
            set
            {
                if (IsReference)
                    Worksheet.Internals.ColumnsCollection[ColumnNumber()].IsHidden = value;
                else
                    _isHidden = value;
            }
        }

        public Int32 OutlineLevel
        {
            get { return IsReference ? Worksheet.Internals.ColumnsCollection[ColumnNumber()].OutlineLevel : _outlineLevel; }
            set
            {
                if (value < 0 || value > 8)
                    throw new ArgumentOutOfRangeException("value", "Outline level must be between 0 and 8.");

                if (IsReference)
                    Worksheet.Internals.ColumnsCollection[ColumnNumber()].OutlineLevel = value;
                else
                {
                    Worksheet.IncrementColumnOutline(value);
                    Worksheet.DecrementColumnOutline(_outlineLevel);
                    _outlineLevel = value;
                }
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

        IXLRangeColumn IXLColumn.CopyTo(IXLCell target)
        {
            using (var asRange = AsRange())
            using (var copy = asRange.CopyTo(target))
                return copy.Column(1);
        }

        IXLRangeColumn IXLColumn.CopyTo(IXLRangeBase target)
        {
            using (var asRange = AsRange())
            using (var copy = asRange.CopyTo(target))
                return copy.Column(1);
        }

        public IXLColumn CopyTo(IXLColumn column)
        {
            column.Clear();
            var newColumn = (XLColumn)column;
            newColumn._width = _width;
            newColumn.InnerStyle = InnerStyle;

            using (var asRange = AsRange())
                asRange.CopyTo(column).Dispose();

            return newColumn;
        }

        public IXLRangeColumn Column(Int32 start, Int32 end)
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
                using (var asRange = AsRange())
                    asRange.Columns(pair.Trim()).ForEach(retVal.Add);
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

        public IXLColumn SetDataType(XLDataType dataType)
        {
            DataType = dataType;
            return this;
        }

        public IXLRangeColumn ColumnUsed(Boolean includeFormats = false)
        {
            return Column(FirstCellUsed(includeFormats), LastCellUsed(includeFormats));
        }

        #endregion IXLColumn Members

        public override XLRange AsRange()
        {
            return Range(1, 1, XLHelper.MaxRowNumber, 1);
        }

        private void WorksheetRangeShiftedColumns(XLRange range, int columnsShifted)
        {
            if (range.RangeAddress.IsValid &&
                RangeAddress.IsValid &&
                range.RangeAddress.FirstAddress.ColumnNumber <= ColumnNumber())
                SetColumnNumber(ColumnNumber() + columnsShifted);
        }

        private void SetColumnNumber(int column)
        {
            if (column <= 0)
                RangeAddress.IsValid = false;
            else
            {
                RangeAddress.IsValid = true;
                RangeAddress.FirstAddress = new XLAddress(Worksheet,
                                                          1,
                                                          column,
                                                          RangeAddress.FirstAddress.FixedRow,
                                                          RangeAddress.FirstAddress.FixedColumn);
                RangeAddress.LastAddress = new XLAddress(Worksheet,
                                                         XLHelper.MaxRowNumber,
                                                         column,
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

        private static double DegreeToRadian(double angle)
        {
            return Math.PI * angle / 180.0;
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
            return false;
        }

        public override Boolean IsEntireColumn()
        {
            return true;
        }
    }
}
