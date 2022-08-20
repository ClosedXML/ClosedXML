using ClosedXML.Excel.Ranges;
using ClosedXML.Excel.Style;
using ClosedXML.Extensions;
using SkiaSharp;
using System;
using System.Collections.Generic;
using System.Linq;

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
        public XLColumn(XLWorksheet worksheet, int column)
            : base(XLRangeAddress.EntireColumn(worksheet, column), worksheet.StyleValue)
        {
            SetColumnNumber(column);

            Width = worksheet.ColumnWidth;
        }

        #endregion Constructor

        public override XLRangeType RangeType => XLRangeType.Column;

        public override IEnumerable<IXLStyle> Styles
        {
            get
            {
                yield return Style;

                var column = ColumnNumber();

                foreach (var cell in Worksheet.Internals.CellsCollection.GetCellsInColumn(column))
                {
                    yield return cell.Style;
                }
            }
        }

        protected override IEnumerable<XLStylizedBase> Children
        {
            get
            {
                var column = ColumnNumber();
                foreach (var cell in Worksheet.Internals.CellsCollection.GetCellsInColumn(column))
                {
                    yield return cell;
                }
            }
        }

        public bool Collapsed { get; set; }

        #region IXLColumn Members

        public double Width { get; set; }

        public void Delete()
        {
            var columnNumber = ColumnNumber();
            Delete(XLShiftDeletedCells.ShiftCellsLeft);
            Worksheet.DeleteColumn(columnNumber);
        }

        public new IXLColumn Clear(XLClearOptions clearOptions = XLClearOptions.All)
        {
            base.Clear(clearOptions);
            return this;
        }

        public IXLCell Cell(int rowNumber)
        {
            return Cell(rowNumber, 1);
        }

        public override IXLCells Cells(string cellsInColumn)
        {
            var retVal = new XLCells(false, XLCellsUsedOptions.All);
            var rangePairs = cellsInColumn.Split(',');
            foreach (var pair in rangePairs)
            {
                retVal.Add(Range(pair.Trim()).RangeAddress);
            }

            return retVal;
        }

        public override IXLCells Cells()
        {
            return Cells(true, XLCellsUsedOptions.All);
        }

        public override IXLCells Cells(bool usedCellsOnly)
        {
            if (usedCellsOnly)
            {
                return Cells(true, XLCellsUsedOptions.AllContents);
            }
            else
            {
                return Cells(FirstCellUsed().Address.RowNumber, LastCellUsed().Address.RowNumber);
            }
        }

        public IXLCells Cells(int firstRow, int lastRow)
        {
            return Cells(firstRow + ":" + lastRow);
        }

        public new IXLColumns InsertColumnsAfter(int numberOfColumns)
        {
            var columnNum = ColumnNumber();
            Worksheet.Internals.ColumnsCollection.ShiftColumnsRight(columnNum + 1, numberOfColumns);
            Worksheet.Column(columnNum).InsertColumnsAfterVoid(true, numberOfColumns);
            var newColumns = Worksheet.Columns(columnNum + 1, columnNum + numberOfColumns);
            CopyColumns(newColumns);
            return newColumns;
        }

        public new IXLColumns InsertColumnsBefore(int numberOfColumns)
        {
            var columnNum = ColumnNumber();
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

        public IXLColumn AdjustToContents(int startRow)
        {
            return AdjustToContents(startRow, XLHelper.MaxRowNumber);
        }

        public IXLColumn AdjustToContents(int startRow, int endRow)
        {
            return AdjustToContents(startRow, endRow, 0, double.MaxValue);
        }

        public IXLColumn AdjustToContents(double minWidth, double maxWidth)
        {
            return AdjustToContents(1, XLHelper.MaxRowNumber, minWidth, maxWidth);
        }

        public IXLColumn AdjustToContents(int startRow, double minWidth, double maxWidth)
        {
            return AdjustToContents(startRow, XLHelper.MaxRowNumber, minWidth, maxWidth);
        }

        public IXLColumn AdjustToContents(int startRow, int endRow, double minWidth, double maxWidth)
        {
            var fontCache = new Dictionary<IXLFontBase, SKFont>();

            var colMaxWidth = minWidth;

            var autoFilterRows = new List<int>();
            if (Worksheet.AutoFilter != null && Worksheet.AutoFilter.Range != null)
            {
                autoFilterRows.Add(Worksheet.AutoFilter.Range.FirstRow().RowNumber());
            }

            autoFilterRows.AddRange(Worksheet.Tables.Where(t =>
                    t.AutoFilter != null
                    && t.AutoFilter.Range != null
                    && !autoFilterRows.Contains(t.AutoFilter.Range.FirstRow().RowNumber()))
                .Select(t => t.AutoFilter.Range.FirstRow().RowNumber()));

            XLStyle cellStyle = null;
            foreach (var c in Column(startRow, endRow).CellsUsed().Cast<XLCell>())
            {
                if (c.IsMerged())
                {
                    continue;
                }

                if (cellStyle == null || cellStyle.Value != c.StyleValue)
                {
                    cellStyle = c.Style as XLStyle;
                }

                double thisWidthMax = 0;
                var textRotation = cellStyle.Alignment.TextRotation;
                if (c.HasRichText || textRotation != 0 || c.InnerText.Contains(XLConstants.NewLine))
                {
                    var kpList = new List<KeyValuePair<IXLFontBase, string>>();

                    #region if (c.HasRichText)

                    if (c.HasRichText)
                    {
                        foreach (var rt in c.GetRichText())
                        {
                            var formattedString = rt.Text;
                            var arr = formattedString.Split(new[] { XLConstants.NewLine }, StringSplitOptions.None);
                            var arrCount = arr.Length;
                            for (var i = 0; i < arrCount; i++)
                            {
                                var s = arr[i];
                                if (i < arrCount - 1)
                                {
                                    s += XLConstants.NewLine;
                                }

                                kpList.Add(new KeyValuePair<IXLFontBase, string>(rt, s));
                            }
                        }
                    }
                    else
                    {
                        var formattedString = c.GetFormattedString();
                        var arr = formattedString.Split(new[] { XLConstants.NewLine }, StringSplitOptions.None);
                        var arrCount = arr.Length;
                        for (var i = 0; i < arrCount; i++)
                        {
                            var s = arr[i];
                            if (i < arrCount - 1)
                            {
                                s += XLConstants.NewLine;
                            }

                            kpList.Add(new KeyValuePair<IXLFontBase, string>(cellStyle.Font, s));
                        }
                    }

                    #endregion if (c.HasRichText)

                    #region foreach (var kp in kpList)

                    double runningWidth = 0;
                    var rotated = false;
                    double maxLineWidth = 0;
                    var lineCount = 1;
                    foreach (var kp in kpList)
                    {
                        var f = kp.Key;
                        var formattedString = kp.Value;

                        var newLinePosition = formattedString.IndexOf(XLConstants.NewLine);
                        if (textRotation == 0)
                        {
                            #region if (newLinePosition >= 0)

                            if (newLinePosition >= 0)
                            {
                                if (newLinePosition > 0)
                                {
                                    runningWidth += f.GetWidth(formattedString.Substring(0, newLinePosition), fontCache);
                                }

                                if (runningWidth > thisWidthMax)
                                {
                                    thisWidthMax = runningWidth;
                                }

                                runningWidth = newLinePosition < formattedString.Length - 2
                                                   ? f.GetWidth(formattedString.Substring(newLinePosition + 2), fontCache)
                                                   : 0;
                            }
                            else
                            {
                                runningWidth += f.GetWidth(formattedString, fontCache);
                            }

                            #endregion if (newLinePosition >= 0)
                        }
                        else
                        {
                            #region if (textRotation == 255)

                            if (textRotation == 255)
                            {
                                if (runningWidth <= 0)
                                {
                                    runningWidth = f.GetWidth("X", fontCache);
                                }

                                if (newLinePosition >= 0)
                                {
                                    runningWidth += f.GetWidth("X", fontCache);
                                }
                            }
                            else
                            {
                                rotated = true;
                                var vWidth = f.GetWidth("X", fontCache);
                                if (vWidth > maxLineWidth)
                                {
                                    maxLineWidth = vWidth;
                                }

                                if (newLinePosition >= 0)
                                {
                                    lineCount++;

                                    if (newLinePosition > 0)
                                    {
                                        runningWidth += f.GetWidth(formattedString.Substring(0, newLinePosition), fontCache);
                                    }

                                    if (runningWidth > thisWidthMax)
                                    {
                                        thisWidthMax = runningWidth;
                                    }

                                    runningWidth = newLinePosition < formattedString.Length - 2
                                                       ? f.GetWidth(formattedString.Substring(newLinePosition + 2), fontCache)
                                                       : 0;
                                }
                                else
                                {
                                    runningWidth += f.GetWidth(formattedString, fontCache);
                                }
                            }

                            #endregion if (textRotation == 255)
                        }
                    }

                    #endregion foreach (var kp in kpList)

                    if (runningWidth > thisWidthMax)
                    {
                        thisWidthMax = runningWidth;
                    }

                    #region if (rotated)

                    if (rotated)
                    {
                        int rotation;
                        if (textRotation == 90 || textRotation == 180 || textRotation == 255)
                        {
                            rotation = 90;
                        }
                        else
                        {
                            rotation = textRotation % 90;
                        }

                        var r = DegreeToRadian(rotation);

                        thisWidthMax = (thisWidthMax * Math.Cos(r)) + (maxLineWidth * lineCount);
                    }

                    #endregion if (rotated)
                }
                else
                {
                    thisWidthMax = cellStyle.Font.GetWidth(c.GetFormattedString(), fontCache);
                }

                if (autoFilterRows.Contains(c.Address.RowNumber))
                {
                    // Allow room for arrow icon in auto filter
                    thisWidthMax += 2.7148;
                }

                if (thisWidthMax >= maxWidth)
                {
                    colMaxWidth = maxWidth;
                    break;
                }

                if (thisWidthMax > colMaxWidth)
                {
                    colMaxWidth = thisWidthMax + 1;
                }
            }

            if (colMaxWidth <= 0)
            {
                colMaxWidth = Worksheet.ColumnWidth;
            }

            Width = Math.Round(colMaxWidth, 1);

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

        public bool IsHidden { get; set; }

        public int OutlineLevel
        {
            get { return _outlineLevel; }
            set
            {
                if (value < 0 || value > 8)
                {
                    throw new ArgumentOutOfRangeException("value", "Outline level must be between 0 and 8.");
                }

                Worksheet.IncrementColumnOutline(value);
                Worksheet.DecrementColumnOutline(_outlineLevel);
                _outlineLevel = value;
            }
        }

        public IXLColumn Group()
        {
            return Group(false);
        }

        public IXLColumn Group(bool collapse)
        {
            if (OutlineLevel < 8)
            {
                OutlineLevel += 1;
            }

            Collapsed = collapse;
            return this;
        }

        public IXLColumn Group(int outlineLevel)
        {
            return Group(outlineLevel, false);
        }

        public IXLColumn Group(int outlineLevel, bool collapse)
        {
            OutlineLevel = outlineLevel;
            Collapsed = collapse;
            return this;
        }

        public IXLColumn Ungroup()
        {
            return Ungroup(false);
        }

        public IXLColumn Ungroup(bool ungroupFromAll)
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

        public int CellCount()
        {
            return RangeAddress.LastAddress.ColumnNumber - RangeAddress.FirstAddress.ColumnNumber + 1;
        }

        public IXLColumn Sort(XLSortOrder sortOrder = XLSortOrder.Ascending, bool matchCase = false,
                              bool ignoreBlanks = true)
        {
            Sort(1, sortOrder, matchCase, ignoreBlanks);
            return this;
        }

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

        public IXLRangeColumn Column(int start, int end)
        {
            return Range(start, 1, end, 1).Column(1);
        }

        public IXLRangeColumn Column(IXLCell start, IXLCell end)
        {
            return Column(start.Address.RowNumber, end.Address.RowNumber);
        }

        public IXLRangeColumns Columns(string columns)
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

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        public IXLRangeColumn ColumnUsed(bool includeFormats)
        {
            return ColumnUsed(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents);
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

        public override XLRange Range(string rangeAddressStr)
        {
            string rangeAddressToUse;
            if (rangeAddressStr.Contains(':') || rangeAddressStr.Contains('-'))
            {
                if (rangeAddressStr.Contains('-'))
                {
                    rangeAddressStr = rangeAddressStr.Replace('-', ':');
                }

                var arrRange = rangeAddressStr.Split(':');
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

        private static double DegreeToRadian(double angle)
        {
            return Math.PI * angle / 180.0;
        }

        private XLColumn ColumnShift(int columnsToShift)
        {
            return Worksheet.Column(ColumnNumber() + columnsToShift);
        }

        #region XLColumn Left

        IXLColumn IXLColumn.ColumnLeft()
        {
            return ColumnLeft();
        }

        IXLColumn IXLColumn.ColumnLeft(int step)
        {
            return ColumnLeft(step);
        }

        public XLColumn ColumnLeft()
        {
            return ColumnLeft(1);
        }

        public XLColumn ColumnLeft(int step)
        {
            return ColumnShift(step * -1);
        }

        #endregion XLColumn Left

        #region XLColumn Right

        IXLColumn IXLColumn.ColumnRight()
        {
            return ColumnRight();
        }

        IXLColumn IXLColumn.ColumnRight(int step)
        {
            return ColumnRight(step);
        }

        public XLColumn ColumnRight()
        {
            return ColumnRight(1);
        }

        public XLColumn ColumnRight(int step)
        {
            return ColumnShift(step);
        }

        #endregion XLColumn Right

        public override bool IsEmpty()
        {
            return IsEmpty(XLCellsUsedOptions.AllContents);
        }

        public override bool IsEmpty(XLCellsUsedOptions options)
        {
            if (options.HasFlag(XLCellsUsedOptions.NormalFormats) &&
                !StyleValue.Equals(Worksheet.StyleValue))
            {
                return false;
            }

            return base.IsEmpty(options);
        }

        public override bool IsEntireRow()
        {
            return false;
        }

        public override bool IsEntireColumn()
        {
            return true;
        }
    }
}