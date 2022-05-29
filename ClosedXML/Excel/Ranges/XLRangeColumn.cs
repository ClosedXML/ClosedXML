using System;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLRangeColumn : XLRangeBase, IXLRangeColumn
    {
        #region Constructor

        /// <summary>
        /// The direct constructor should only be used in <see cref="XLWorksheet.RangeFactory"/>.
        /// </summary>
        public XLRangeColumn(XLRangeParameters rangeParameters)
            : base(rangeParameters.RangeAddress, (rangeParameters.DefaultStyle as XLStyle).Value)
        {
        }

        #endregion Constructor

        #region IXLRangeColumn Members

        IXLCell IXLRangeColumn.Cell(int rowNumber)
        {
            return Cell(rowNumber);
        }

        public override IXLCells Cells(string cellsInColumn)
        {
            var retVal = new XLCells(false, XLCellsUsedOptions.AllContents);
            var rangePairs = cellsInColumn.Split(',');
            foreach (string pair in rangePairs)
                retVal.Add(Range(pair.Trim()).RangeAddress);
            return retVal;
        }

        public IXLCells Cells(int firstRow, int lastRow)
        {
            return Cells(firstRow + ":" + lastRow);
        }

        public void Delete()
        {
            Delete(true);
        }

        internal void Delete(Boolean deleteTableField)
        {
            if (deleteTableField && IsTableColumn())
            {
                var table = Table as XLTable;
                var firstCellValue = Cell(1).Value.ToString();
                if (!table.FieldNames.ContainsKey(firstCellValue))
                    throw new ArgumentException(string.Format("Field {0} not found.", firstCellValue));

                var field = table.Fields.Cast<XLTableField>().Single(f => f.Name == firstCellValue);
                field.Delete(false);
            }

            Delete(XLShiftDeletedCells.ShiftCellsLeft);
        }

        public IXLCells InsertCellsAbove(int numberOfRows)
        {
            return InsertCellsAbove(numberOfRows, false);
        }

        public IXLCells InsertCellsAbove(int numberOfRows, bool expandRange)
        {
            return InsertRowsAbove(numberOfRows, expandRange).Cells();
        }

        public IXLCells InsertCellsBelow(int numberOfRows)
        {
            return InsertCellsBelow(numberOfRows, true);
        }

        public IXLCells InsertCellsBelow(int numberOfRows, bool expandRange)
        {
            return InsertRowsBelow(numberOfRows, expandRange).Cells();
        }

        public int CellCount()
        {
            return RangeAddress.LastAddress.RowNumber - RangeAddress.FirstAddress.RowNumber + 1;
        }

        public IXLRangeColumn Sort(XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true)
        {
            base.Sort(1, sortOrder, matchCase, ignoreBlanks);
            return this;
        }

        public new IXLRangeColumn CopyTo(IXLCell target)
        {
            base.CopyTo(target);

            int lastRowNumber = target.Address.RowNumber + RowCount() - 1;
            if (lastRowNumber > XLHelper.MaxRowNumber)
                lastRowNumber = XLHelper.MaxRowNumber;
            int lastColumnNumber = target.Address.ColumnNumber + ColumnCount() - 1;
            if (lastColumnNumber > XLHelper.MaxColumnNumber)
                lastColumnNumber = XLHelper.MaxColumnNumber;

            return target.Worksheet.Range(
                target.Address.RowNumber,
                target.Address.ColumnNumber,
                lastRowNumber,
                lastColumnNumber)
                .Column(1);
        }

        public new IXLRangeColumn CopyTo(IXLRangeBase target)
        {
            base.CopyTo(target);

            int lastRowNumber = target.RangeAddress.FirstAddress.RowNumber + RowCount() - 1;
            if (lastRowNumber > XLHelper.MaxRowNumber)
                lastRowNumber = XLHelper.MaxRowNumber;
            int lastColumnNumber = target.RangeAddress.FirstAddress.ColumnNumber + ColumnCount() - 1;
            if (lastColumnNumber > XLHelper.MaxColumnNumber)
                lastColumnNumber = XLHelper.MaxColumnNumber;

            return target.Worksheet.Range(
                target.RangeAddress.FirstAddress.RowNumber,
                target.RangeAddress.FirstAddress.ColumnNumber,
                lastRowNumber,
                lastColumnNumber)
                .Column(1);
        }

        public IXLRangeColumn Column(int start, int end)
        {
            return Range(start, end).FirstColumn();
        }

        public IXLRangeColumn Column(IXLCell start, IXLCell end)
        {
            return Column(start.Address.RowNumber, end.Address.RowNumber);
        }

        public IXLRangeColumns Columns(string columns)
        {
            var retVal = new XLRangeColumns();
            var rowPairs = columns.Split(',');
            foreach (string trimmedPair in rowPairs.Select(pair => pair.Trim()))
            {
                string firstRow;
                string lastRow;
                if (trimmedPair.Contains(':') || trimmedPair.Contains('-'))
                {
                    var rowRange = trimmedPair.Split(':', '-');

                    firstRow = rowRange[0];
                    lastRow = rowRange[1];
                }
                else
                {
                    firstRow = trimmedPair;
                    lastRow = trimmedPair;
                }

                retVal.Add(Range(firstRow, lastRow).FirstColumn());
            }

            return retVal;
        }

        public IXLRangeColumn SetDataType(XLDataType dataType)
        {
            DataType = dataType;
            return this;
        }

        public IXLColumn WorksheetColumn()
        {
            return Worksheet.Column(RangeAddress.FirstAddress.ColumnNumber);
        }

        #endregion IXLRangeColumn Members

        public override XLRangeType RangeType
        {
            get { return XLRangeType.RangeColumn; }
        }

        public XLCell Cell(int row)
        {
            return Cell(row, 1);
        }

        internal override void WorksheetRangeShiftedColumns(XLRange range, int columnsShifted)
        {
            RangeAddress = (XLRangeAddress)ShiftColumns(RangeAddress, range, columnsShifted);
        }

        internal override void WorksheetRangeShiftedRows(XLRange range, int rowsShifted)
        {
            RangeAddress = (XLRangeAddress)ShiftRows(RangeAddress, range, rowsShifted);
        }

        public XLRange Range(int firstRow, int lastRow)
        {
            return Range(firstRow, 1, lastRow, 1);
        }

        public override XLRange Range(string rangeAddressStr)
        {
            string rangeAddressToUse;
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

        public int CompareTo(XLRangeColumn otherColumn, IXLSortElements rowsToSort)
        {
            foreach (IXLSortElement e in rowsToSort)
            {
                var thisCell = Cell(e.ElementNumber);
                var otherCell = otherColumn.Cell(e.ElementNumber);
                int comparison;
                bool thisCellIsBlank = thisCell.IsEmpty();
                bool otherCellIsBlank = otherCell.IsEmpty();
                if (e.IgnoreBlanks && (thisCellIsBlank || otherCellIsBlank))
                {
                    if (thisCellIsBlank && otherCellIsBlank)
                        comparison = 0;
                    else
                    {
                        if (thisCellIsBlank)
                            comparison = e.SortOrder == XLSortOrder.Ascending ? 1 : -1;
                        else
                            comparison = e.SortOrder == XLSortOrder.Ascending ? -1 : 1;
                    }
                }
                else
                {
                    if (thisCell.DataType == otherCell.DataType)
                    {
                        if (thisCell.DataType == XLDataType.Text)
                        {
                            comparison = e.MatchCase
                                             ? thisCell.InnerText.CompareTo(otherCell.InnerText)
                                             : String.Compare(thisCell.InnerText, otherCell.InnerText, true);
                        }
                        else if (thisCell.DataType == XLDataType.TimeSpan)
                            comparison = thisCell.GetTimeSpan().CompareTo(otherCell.GetTimeSpan());
                        else
                            comparison = Double.Parse(thisCell.InnerText, XLHelper.NumberStyle, XLHelper.ParseCulture).CompareTo(Double.Parse(otherCell.InnerText, XLHelper.NumberStyle, XLHelper.ParseCulture));
                    }
                    else if (e.MatchCase)
                        comparison = String.Compare(thisCell.GetString(), otherCell.GetString(), true);
                    else
                        comparison = thisCell.GetString().CompareTo(otherCell.GetString());
                }

                if (comparison != 0)
                    return e.SortOrder == XLSortOrder.Ascending ? comparison : comparison * -1;
            }

            return 0;
        }

        private XLRangeColumn ColumnShift(Int32 columnsToShift)
        {
            Int32 columnNumber = ColumnNumber() + columnsToShift;
            return Worksheet.Range(
                RangeAddress.FirstAddress.RowNumber,
                columnNumber,
                RangeAddress.LastAddress.RowNumber,
                columnNumber).FirstColumn();
        }

        #region XLRangeColumn Left

        IXLRangeColumn IXLRangeColumn.ColumnLeft()
        {
            return ColumnLeft();
        }

        IXLRangeColumn IXLRangeColumn.ColumnLeft(Int32 step)
        {
            return ColumnLeft(step);
        }

        public XLRangeColumn ColumnLeft()
        {
            return ColumnLeft(1);
        }

        public XLRangeColumn ColumnLeft(Int32 step)
        {
            return ColumnShift(step * -1);
        }

        #endregion XLRangeColumn Left

        #region XLRangeColumn Right

        IXLRangeColumn IXLRangeColumn.ColumnRight()
        {
            return ColumnRight();
        }

        IXLRangeColumn IXLRangeColumn.ColumnRight(Int32 step)
        {
            return ColumnRight(step);
        }

        public XLRangeColumn ColumnRight()
        {
            return ColumnRight(1);
        }

        public XLRangeColumn ColumnRight(Int32 step)
        {
            return ColumnShift(step);
        }

        #endregion XLRangeColumn Right

        public IXLTable AsTable()
        {
            if (IsTableColumn())
                throw new InvalidOperationException("This column is already part of a table.");

            return AsRange().AsTable();
        }

        public IXLTable AsTable(string name)
        {
            if (IsTableColumn())
                throw new InvalidOperationException("This column is already part of a table.");

            return AsRange().AsTable(name);
        }

        public IXLTable CreateTable()
        {
            if (IsTableColumn())
                throw new InvalidOperationException("This column is already part of a table.");

            return AsRange().CreateTable();
        }

        public IXLTable CreateTable(string name)
        {
            if (IsTableColumn())
                throw new InvalidOperationException("This column is already part of a table.");

            return AsRange().CreateTable(name);
        }

        public new IXLRangeColumn Clear(XLClearOptions clearOptions = XLClearOptions.All)
        {
            base.Clear(clearOptions);
            return this;
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        public IXLRangeColumn ColumnUsed(Boolean includeFormats)
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

        internal IXLTable Table { get; set; }

        public Boolean IsTableColumn()
        {
            return Table != null;
        }
    }
}
