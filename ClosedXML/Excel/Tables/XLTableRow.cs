﻿using System;

namespace ClosedXML.Excel
{
    internal class XLTableRow : XLRangeRow, IXLTableRow
    {
        private readonly XLTableRange _tableRange;

        public XLTableRow(XLTableRange tableRange, XLRangeRow rangeRow)
            : base(rangeRow.RangeParameters, false)
        {
            _tableRange = tableRange;
        }

        #region IXLTableRow Members

        public IXLCell Field(Int32 index)
        {
            return Cell(index + 1);
        }

        public IXLCell Field(String name)
        {
            Int32 fieldIndex = _tableRange.Table.GetFieldIndex(name);
            return Cell(fieldIndex + 1);
        }

        public new IXLTableRow Sort()
        {
            return SortLeftToRight();
        }

        public new IXLTableRow SortLeftToRight(XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true)
        {
            base.SortLeftToRight(sortOrder, matchCase, ignoreBlanks);
            return this;
        }

        #endregion

        private XLTableRow RowShift(Int32 rowsToShift)
        {
            return _tableRange.Row(RowNumber() + rowsToShift);
        }

        #region XLTableRow Above

        IXLTableRow IXLTableRow.RowAbove()
        {
            return RowAbove();
        }

        IXLTableRow IXLTableRow.RowAbove(Int32 step)
        {
            return RowAbove(step);
        }

        public new XLTableRow RowAbove()
        {
            return RowAbove(1);
        }

        public new XLTableRow RowAbove(Int32 step)
        {
            return RowShift(step * -1);
        }

        #endregion

        #region XLTableRow Below

        IXLTableRow IXLTableRow.RowBelow()
        {
            return RowBelow();
        }

        IXLTableRow IXLTableRow.RowBelow(Int32 step)
        {
            return RowBelow(step);
        }

        public new XLTableRow RowBelow()
        {
            return RowBelow(1);
        }

        public new XLTableRow RowBelow(Int32 step)
        {
            return RowShift(step);
        }

        #endregion

        public new IXLTableRow Clear(XLClearOptions clearOptions = XLClearOptions.ContentsAndFormats)
        {
            base.Clear(clearOptions);
            return this;
        }

        public new IXLTableRows InsertRowsAbove(int numberOfRows)
        {
            return XLHelper.InsertRowsWithoutEvents(base.InsertRowsAbove, _tableRange, numberOfRows, !_tableRange.Table.ShowTotalsRow);
        }
        public new IXLTableRows InsertRowsBelow(int numberOfRows)
        {
            return XLHelper.InsertRowsWithoutEvents(base.InsertRowsBelow, _tableRange, numberOfRows, !_tableRange.Table.ShowTotalsRow);
        }

        public new void Delete()
        {
            Delete(XLShiftDeletedCells.ShiftCellsUp);
            _tableRange.Table.ExpandTableRows(-1);
        }
    }
}