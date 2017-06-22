using System;

namespace ClosedXML.Excel
{
    using System.Linq;

    internal class XLTableRange : XLRange, IXLTableRange
    {
        private readonly XLTable _table;
        private readonly XLRange _range;
        public XLTableRange(XLRange range, XLTable table):base(range.RangeParameters)
        {
            _table = table;
            _range = range;
        }

        IXLTableRow IXLTableRange.FirstRow(Func<IXLTableRow, Boolean> predicate)
        {
            return FirstRow(predicate);
        }
        public XLTableRow FirstRow(Func<IXLTableRow, Boolean> predicate = null)
        {
            if (predicate == null)
                return new XLTableRow(this, (_range.FirstRow()));

            Int32 rowCount = _range.RowCount();

            for (Int32 ro = 1; ro <= rowCount; ro++)
            {
                var row = new XLTableRow(this, (_range.Row(ro)));
                if (predicate(row)) return row;

                row.Dispose();
            }

            return null;
        }

        IXLTableRow IXLTableRange.FirstRowUsed(Func<IXLTableRow, Boolean> predicate)
        {
            return FirstRowUsed(false, predicate);
        }
        public XLTableRow FirstRowUsed(Func<IXLTableRow, Boolean> predicate = null)
        {
            return FirstRowUsed(false, predicate);
        }

        IXLTableRow IXLTableRange.FirstRowUsed(Boolean includeFormats, Func<IXLTableRow, Boolean> predicate)
        {
            return FirstRowUsed(includeFormats, predicate);
        }
        public XLTableRow FirstRowUsed(Boolean includeFormats, Func<IXLTableRow, Boolean> predicate = null)
        {
            if (predicate == null)
                return new XLTableRow(this, (_range.FirstRowUsed(includeFormats)));

            Int32 rowCount = _range.RowCount();

            for (Int32 ro = 1; ro <= rowCount; ro++)
            {
                var row = new XLTableRow(this, (_range.Row(ro)));

                if (!row.IsEmpty(includeFormats) && predicate(row))
                    return row;
                row.Dispose();
            }

            return null;
        }


        IXLTableRow IXLTableRange.LastRow(Func<IXLTableRow, Boolean> predicate)
        {
            return LastRow(predicate);
        }
        public XLTableRow LastRow(Func<IXLTableRow, Boolean> predicate = null)
        {
            if (predicate == null)
                return new XLTableRow(this, (_range.LastRow()));

            Int32 rowCount = _range.RowCount();

            for (Int32 ro = rowCount; ro >= 1; ro--)
            {
                var row = new XLTableRow(this, (_range.Row(ro)));
                if (predicate(row)) return row;

                row.Dispose();
            }
            return null;
        }

        IXLTableRow IXLTableRange.LastRowUsed(Func<IXLTableRow, Boolean> predicate)
        {
            return LastRowUsed(false, predicate);
        }
        public XLTableRow LastRowUsed(Func<IXLTableRow, Boolean> predicate = null)
        {
            return LastRowUsed(false, predicate);
        }

        IXLTableRow IXLTableRange.LastRowUsed(Boolean includeFormats, Func<IXLTableRow, Boolean> predicate)
        {
            return LastRowUsed(includeFormats, predicate);
        }
        public XLTableRow LastRowUsed(Boolean includeFormats, Func<IXLTableRow, Boolean> predicate = null)
        {
            if (predicate == null)
                return new XLTableRow(this, (_range.LastRowUsed(includeFormats)));

            Int32 rowCount = _range.RowCount();

            for (Int32 ro = rowCount; ro >= 1; ro--)
            {
                var row = new XLTableRow(this, (_range.Row(ro)));

                if (!row.IsEmpty(includeFormats) && predicate(row))
                    return row;
                row.Dispose();
            }

            return null;
        }

        IXLTableRow IXLTableRange.Row(int row)
        {
            return Row(row);
        }
        public new XLTableRow Row(int row)
        {
            if (row <= 0 || row > XLHelper.MaxRowNumber)
            {
                throw new IndexOutOfRangeException(String.Format("Row number must be between 1 and {0}",
                                                                 XLHelper.MaxRowNumber));
            }

            return new XLTableRow(this, base.Row(row));
        }

        public IXLTableRows Rows(Func<IXLTableRow, Boolean> predicate = null)
        {
            var retVal = new XLTableRows(Worksheet.Style);
            Int32 rowCount = _range.RowCount();

            for (int r = 1; r <= rowCount; r++)
            {
                var row = Row(r);
                if (predicate == null || predicate(row))
                    retVal.Add(row);
                else
                    row.Dispose();
            }
            return retVal;
        }

        public new IXLTableRows Rows(int firstRow, int lastRow)
        {
            var retVal = new XLTableRows(Worksheet.Style);

            for (int ro = firstRow; ro <= lastRow; ro++)
                retVal.Add(Row(ro));
            return retVal;
        }

        public new IXLTableRows Rows(string rows)
        {
            var retVal = new XLTableRows(Worksheet.Style);
            var rowPairs = rows.Split(',');
            foreach (string tPair in rowPairs.Select(pair => pair.Trim()))
            {
                String firstRow;
                String lastRow;
                if (tPair.Contains(':') || tPair.Contains('-'))
                {
                    var rowRange = XLHelper.SplitRange(tPair);

                    firstRow = rowRange[0];
                    lastRow = rowRange[1];
                }
                else
                {
                    firstRow = tPair;
                    lastRow = tPair;
                }
                foreach (IXLTableRow row in Rows(Int32.Parse(firstRow), Int32.Parse(lastRow)))
                    retVal.Add(row);
            }
            return retVal;
        }

        IXLTableRows IXLTableRange.RowsUsed(Boolean includeFormats, Func<IXLTableRow, Boolean> predicate)
        {
            return RowsUsed(includeFormats, predicate);
        }
        public IXLTableRows RowsUsed(Boolean includeFormats, Func<IXLTableRow, Boolean> predicate = null)
        {
            var rows = new XLTableRows(Worksheet.Style);
            Int32 rowCount = RowCount();

            for (Int32 ro = 1; ro <= rowCount; ro++)
            {
                var row = Row(ro);

                if (!row.IsEmpty(includeFormats) && (predicate == null || predicate(row)))
                    rows.Add(row);
                else
                    row.Dispose();
            }
            return rows;
        }

        IXLTableRows IXLTableRange.RowsUsed(Func<IXLTableRow, Boolean> predicate)
        {
            return RowsUsed(predicate);
        }
        public IXLTableRows RowsUsed(Func<IXLTableRow, Boolean> predicate = null)
        {
            return RowsUsed(false, predicate);
        }

        IXLTable IXLTableRange.Table { get { return _table; } }
        public XLTable Table { get { return _table; } }

        public new IXLTableRows InsertRowsAbove(int numberOfRows)
        {
            return XLHelper.InsertRowsWithoutEvents(base.InsertRowsAbove, this, numberOfRows, !Table.ShowTotalsRow );
        }
        public new IXLTableRows InsertRowsBelow(int numberOfRows)
        {
            return XLHelper.InsertRowsWithoutEvents(base.InsertRowsBelow, this, numberOfRows, !Table.ShowTotalsRow);
        }


        public new IXLRangeColumn Column(String column)
        {
            if (XLHelper.IsValidColumn(column))
            {
                Int32 coNum = XLHelper.GetColumnNumberFromLetter(column);
                return coNum > ColumnCount() ? Column(_table.GetFieldIndex(column) + 1) : Column(coNum);
            }

            return Column(_table.GetFieldIndex(column) + 1);
        }
    }
}