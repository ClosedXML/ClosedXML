using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace ClosedXML.Excel
{
    internal class XLColumn: XLRangeBase, IXLColumn
    {
        public XLColumn(Int32 column, XLColumnParameters xlColumnParameters)
        {
            SetColumnNumber(column);
            Worksheet = xlColumnParameters.Worksheet;
            

            this.IsReference = xlColumnParameters.IsReference;
            if (IsReference)
            {
                Worksheet.RangeShiftedColumns += new RangeShiftedColumnsDelegate(Worksheet_RangeShiftedColumns);
            }
            else
            {
                this.style = new XLStyle(this, xlColumnParameters.DefaultStyle);
                this.width = xlColumnParameters.Worksheet.ColumnWidth;
            }
        }

        void Worksheet_RangeShiftedColumns(XLRange range, int columnsShifted)
        {
            if (range.FirstAddressInSheet.ColumnNumber <= this.ColumnNumber())
                SetColumnNumber(this.ColumnNumber() + columnsShifted);
        }

        private void SetColumnNumber(Int32 column)
        {
            FirstAddressInSheet = new XLAddress(1, column);
            LastAddressInSheet = new XLAddress(XLWorksheet.MaxNumberOfRows, column);
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
                    return Worksheet.Internals.ColumnsCollection[this.ColumnNumber()].Width;
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
                    Worksheet.Internals.ColumnsCollection[this.ColumnNumber()].Width = value;
                }
                else
                {
                    width = value;
                }
            }
        }

        public void Delete()
        {
            var columnNumber = this.ColumnNumber();
            this.AsRange().Delete(XLShiftDeletedCells.ShiftCellsLeft);
            Worksheet.Internals.ColumnsCollection.Remove(columnNumber);
        }

        public new void Clear()
        {
            var range = this.AsRange();
            range.Clear();
            this.Style = Worksheet.Style;
        }

        public IXLCell Cell(int row)
        {
            return base.Cell(row, 1);
        }

        #endregion

        #region IXLStylized Members

        private IXLStyle style;
        public override IXLStyle Style
        {
            get
            {
                if (IsReference)
                    return Worksheet.Internals.ColumnsCollection[this.ColumnNumber()].Style;
                else
                    return style;
            }
            set
            {
                if (IsReference)
                {
                    Worksheet.Internals.ColumnsCollection[this.ColumnNumber()].Style = value;
                }
                else
                {
                    style = new XLStyle(this, value);

                    foreach (var c in Worksheet.Internals.CellsCollection.Values.Where(c => c.Address.ColumnNumber == this.ColumnNumber()))
                    {
                        c.Style = value;
                    }

                    var maxRow = 0;
                    if (Worksheet.Internals.RowsCollection.Count > 0)
                        maxRow = Worksheet.Internals.RowsCollection.Keys.Max();

                    for (var ro = 1; ro <= maxRow; ro++)
                    {
                        Worksheet.Cell(ro, this.ColumnNumber()).Style = value;
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

                var co = this.ColumnNumber();

                foreach (var c in Worksheet.Internals.CellsCollection.Values.Where(c => c.Address.ColumnNumber == co))
                {
                    yield return c.Style;
                }

                var maxRow = 0;
                if (Worksheet.Internals.RowsCollection.Count > 0)
                    maxRow = Worksheet.Internals.RowsCollection.Keys.Max();

                for (var ro = 1; ro <= maxRow; ro++)
                {
                    yield return Worksheet.Cell(ro, co).Style;
                }

                UpdatingStyle = false;
            }
        }

        public override Boolean UpdatingStyle { get; set; }

        #endregion

        public Int32 ColumnNumber()
        {
            return this.FirstAddressInSheet.ColumnNumber;
        }
        public String ColumnLetter()
        {
            return this.FirstAddressInSheet.ColumnLetter;
        }

        public new void InsertColumnsAfter( Int32 numberOfColumns)
        {
            var columnNum = this.ColumnNumber();
            this.Worksheet.Internals.ColumnsCollection.ShiftColumnsRight(columnNum + 1, numberOfColumns);
            XLRange range = (XLRange)this.Worksheet.Column(columnNum).AsRange();
            range.InsertColumnsAfter(numberOfColumns, true);
        }
        public new void InsertColumnsBefore( Int32 numberOfColumns)
        {
            var columnNum = this.ColumnNumber();
            this.Worksheet.Internals.ColumnsCollection.ShiftColumnsRight(columnNum, numberOfColumns);
            // We can't use this.AsRange() because we've shifted the columns
            // and we want to use the old columnNum.
            XLRange range = (XLRange)this.Worksheet.Column(columnNum).AsRange(); 
            range.InsertColumnsBefore(numberOfColumns, true);
        }

        public override IXLRange AsRange()
        {
            return Range(1, 1, XLWorksheet.MaxNumberOfRows, 1);
        }
    }
}
