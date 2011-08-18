using System;

namespace ClosedXML.Excel
{
    internal class XLTableRow : XLRangeRow, IXLTableRow
    {
        private readonly XLTable _table;

        public XLTableRow(XLTable table, XLRangeRow rangeRow)
            : base(rangeRow.RangeParameters, false)
        {
            _table = table;
        }

        #region IXLTableRow Members

        public IXLCell Field(Int32 index)
        {
            return Cell(index + 1);
        }

        public IXLCell Field(String name)
        {
            Int32 fieldIndex = _table.GetFieldIndex(name);
            return Cell(fieldIndex + 1);
        }

        public new IXLTableRow Sort()
        {
            AsRange().Sort(XLSortOrientation.LeftToRight);
            return this;
        }

        public new IXLTableRow Sort(XLSortOrder sortOrder)
        {
            AsRange().Sort(XLSortOrientation.LeftToRight, sortOrder);
            return this;
        }

        public new IXLTableRow Sort(Boolean matchCase)
        {
            AsRange().Sort(XLSortOrientation.LeftToRight, matchCase);
            return this;
        }

        public new IXLTableRow Sort(XLSortOrder sortOrder, Boolean matchCase)
        {
            AsRange().Sort(XLSortOrientation.LeftToRight, sortOrder, matchCase);
            return this;
        }

        #endregion

        private XLTableRow RowShift(Int32 rowsToShift)
        {
            return _table.Row(RowNumber() + rowsToShift);
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
    }
}