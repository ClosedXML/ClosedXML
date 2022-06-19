namespace ClosedXML.Excel
{
    internal class XLTableRow : XLRangeRow, IXLTableRow
    {
        private readonly XLTableRange _tableRange;

        public XLTableRow(XLTableRange tableRange, XLRangeRow rangeRow)
            : base(new XLRangeParameters(rangeRow.RangeAddress, rangeRow.Style))
        {
            _tableRange = tableRange;
        }

        #region IXLTableRow Members

        public IXLCell Field(int index)
        {
            return Cell(index + 1);
        }

        public IXLCell Field(string name)
        {
            var fieldIndex = _tableRange.Table.GetFieldIndex(name);
            return Cell(fieldIndex + 1);
        }

        public new IXLTableRow Sort()
        {
            return SortLeftToRight();
        }

        public new IXLTableRow SortLeftToRight(XLSortOrder sortOrder = XLSortOrder.Ascending, bool matchCase = false, bool ignoreBlanks = true)
        {
            base.SortLeftToRight(sortOrder, matchCase, ignoreBlanks);
            return this;
        }

        #endregion IXLTableRow Members

        private XLTableRow RowShift(int rowsToShift)
        {
            return _tableRange.Row(RowNumber() - _tableRange.FirstRow().RowNumber() + 1 + rowsToShift);
        }

        #region XLTableRow Above

        IXLTableRow IXLTableRow.RowAbove()
        {
            return RowAbove();
        }

        IXLTableRow IXLTableRow.RowAbove(int step)
        {
            return RowAbove(step);
        }

        public new XLTableRow RowAbove()
        {
            return RowAbove(1);
        }

        public new XLTableRow RowAbove(int step)
        {
            return RowShift(step * -1);
        }

        #endregion XLTableRow Above

        #region XLTableRow Below

        IXLTableRow IXLTableRow.RowBelow()
        {
            return RowBelow();
        }

        IXLTableRow IXLTableRow.RowBelow(int step)
        {
            return RowBelow(step);
        }

        public new XLTableRow RowBelow()
        {
            return RowBelow(1);
        }

        public new XLTableRow RowBelow(int step)
        {
            return RowShift(step);
        }

        #endregion XLTableRow Below

        public new IXLTableRow Clear(XLClearOptions clearOptions = XLClearOptions.All)
        {
            base.Clear(clearOptions);
            return this;
        }

        public new IXLTableRows InsertRowsAbove(int numberOfRows)
        {
            return XLHelper.InsertRowsWithoutEvents(InsertRowsAbove, _tableRange, numberOfRows, !_tableRange.Table.ShowTotalsRow);
        }

        public new IXLTableRows InsertRowsBelow(int numberOfRows)
        {
            return XLHelper.InsertRowsWithoutEvents(InsertRowsBelow, _tableRange, numberOfRows, !_tableRange.Table.ShowTotalsRow);
        }

        public new void Delete()
        {
            Delete(XLShiftDeletedCells.ShiftCellsUp);
        }
    }
}
