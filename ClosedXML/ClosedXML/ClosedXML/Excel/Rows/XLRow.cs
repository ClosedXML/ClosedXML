using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace ClosedXML.Excel
{
    internal class XLRow: XLRangeBase, IXLRow
    {
        public XLRow(Int32 row, XLRowParameters xlRowParameters)
            : base(new XLRangeAddress(row, 1, row, XLWorksheet.MaxNumberOfColumns))
        {
            SetRowNumber(row);
            Worksheet = xlRowParameters.Worksheet;

            this.IsReference = xlRowParameters.IsReference;
            if (IsReference)
            {
                Worksheet.RangeShiftedRows += new RangeShiftedRowsDelegate(Worksheet_RangeShiftedRows);
            }
            else
            {
                this.style = new XLStyle(this, xlRowParameters.DefaultStyle);
                this.height = xlRowParameters.Worksheet.RowHeight;
            }
        }

        void Worksheet_RangeShiftedRows(XLRange range, int rowsShifted)
        {
            if (range.RangeAddress.FirstAddress.RowNumber <= this.RowNumber())
                SetRowNumber(this.RowNumber() + rowsShifted);
        }

        void RowsCollection_RowShifted(int startingRow, int rowsShifted)
        {
            if (startingRow <= this.RowNumber())
                SetRowNumber(this.RowNumber() + rowsShifted);
        }

        private void SetRowNumber(Int32 row)
        {
            RangeAddress.FirstAddress = new XLAddress(row, 1);
            RangeAddress.LastAddress = new XLAddress(row, XLWorksheet.MaxNumberOfColumns);
        }

        public Boolean IsReference { get; private set; }

        #region IXLRow Members

        private Double height;
        public Double Height 
        {
            get
            {
                if (IsReference)
                {
                    return Worksheet.Internals.RowsCollection[this.RowNumber()].Height;
                }
                else
                {
                    return height;
                }
            }
            set
            {
                if (IsReference)
                {
                    Worksheet.Internals.RowsCollection[this.RowNumber()].Height = value;
                }
                else
                {
                    height = value;
                }
            }
        }

        public void Delete()
        {
            var rowNumber = this.RowNumber();
            this.AsRange().Delete(XLShiftDeletedCells.ShiftCellsUp);
            Worksheet.Internals.RowsCollection.Remove(rowNumber);
        }

        public Int32 RowNumber()
        {
            return this.RangeAddress.FirstAddress.RowNumber;
        }

        public new void InsertRowsBelow(Int32 numberOfRows)
        {
            var rowNum = this.RowNumber();
            this.Worksheet.Internals.RowsCollection.ShiftRowsDown(rowNum + 1, numberOfRows);
            XLRange range = (XLRange)this.Worksheet.Row(rowNum).AsRange();
            range.InsertRowsBelow(numberOfRows, true);
        }

        public new void InsertRowsAbove(Int32 numberOfRows)
        {
            var rowNum = this.RowNumber();
            this.Worksheet.Internals.RowsCollection.ShiftRowsDown(rowNum, numberOfRows);
            // We can't use this.AsRange() because we've shifted the rows
            // and we want to use the old rowNum.
            XLRange range = (XLRange)this.Worksheet.Row(rowNum).AsRange(); 
            range.InsertRowsAbove(numberOfRows, true);
        }

        public new void Clear()
        {
            var range = this.AsRange();
            range.Clear();
            this.Style = Worksheet.Style;
        }

        public IXLCell Cell(Int32 column)
        {
            return base.Cell(1, column);
        }
        public new IXLCell Cell(String column)
        {
            return base.Cell(1, column);
        }

        public void AdjustToContents()
        {
            Double maxHeight = 0;
            var cellsUsed = CellsUsed();
            foreach (var c in cellsUsed)
            {
                var thisHeight = ((XLFont)c.Style.Font).GetHeight();
                if (thisHeight > maxHeight)
                    maxHeight = thisHeight;
            }
            if (maxHeight > 0)
                Height = maxHeight;
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
                    return Worksheet.Internals.RowsCollection[this.RowNumber()].IsHidden;
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
                    Worksheet.Internals.RowsCollection[this.RowNumber()].IsHidden = value;
                }
                else
                {
                    isHidden = value;
                }
            }
        }

        #endregion

        #region IXLStylized Members

        private IXLStyle style;
        public override IXLStyle Style
        {
            get
            {
                if (IsReference)
                    return Worksheet.Internals.RowsCollection[this.RowNumber()].Style;
                else
                    return style;
            }
            set
            {
                if (IsReference)
                {
                    Worksheet.Internals.RowsCollection[this.RowNumber()].Style = value;
                }
                else
                {
                    style = new XLStyle(this, value);

                    var row = this.RowNumber();
                    foreach (var c in Worksheet.Internals.CellsCollection.Values.Where(c => c.Address.RowNumber == row))
                    {
                        c.Style = value;
                    }

                    var maxColumn = 0;
                    if (Worksheet.Internals.ColumnsCollection.Count > 0)
                        maxColumn = Worksheet.Internals.ColumnsCollection.Keys.Max();


                    for (var co = 1; co <= maxColumn; co++)
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

                var row = this.RowNumber();

                foreach (var c in Worksheet.Internals.CellsCollection.Values.Where(c => c.Address.RowNumber == row))
                {
                    yield return c.Style;
                }
                
                var maxColumn = 0;
                if (Worksheet.Internals.ColumnsCollection.Count > 0)
                    maxColumn = Worksheet.Internals.ColumnsCollection.Keys.Max();

                for (var co = 1; co <= maxColumn; co++)
                {
                    yield return Worksheet.Cell(row, co).Style;
                }

                UpdatingStyle = false;
            }
        }

        public override Boolean UpdatingStyle { get; set; }

        public override IXLRange AsRange()
        {
            return Range(1, 1, 1, XLWorksheet.MaxNumberOfColumns);
        }

        #endregion

        private Boolean collapsed;
        public Boolean Collapsed
        {
            get
            {
                if (IsReference)
                {
                    return Worksheet.Internals.RowsCollection[this.RowNumber()].Collapsed;
                }
                else
                {
                    return collapsed;
                }
            }
            set
            {
                if (IsReference)
                {
                    Worksheet.Internals.RowsCollection[this.RowNumber()].Collapsed = value;
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
            get
            {
                if (IsReference)
                {
                    return Worksheet.Internals.RowsCollection[this.RowNumber()].OutlineLevel;
                }
                else
                {
                    return outlineLevel;
                }
            }
            set
            {
                if (value < 1 || value > 8)
                    throw new ArgumentOutOfRangeException("Outline level must be between 1 and 8.");

                if (IsReference)
                {
                    Worksheet.Internals.RowsCollection[this.RowNumber()].OutlineLevel = value;
                }
                else
                {
                    outlineLevel = value;
                }
            }
        }

        public void Group(Boolean collapse = false)
        {
            if (OutlineLevel < 8)
                OutlineLevel += 1;

            Collapsed = collapse;
        }
        public void Group(Int32 outlineLevel, Boolean collapse = false)
        {
            OutlineLevel = outlineLevel;
            Collapsed = collapse;
        }
        public void Ungroup(Boolean ungroupFromAll = false)
        {
            if (ungroupFromAll)
            {
                OutlineLevel = 0;
            }
            else
            {
                if (OutlineLevel > 0)
                    OutlineLevel -= 1;
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
    }
}
