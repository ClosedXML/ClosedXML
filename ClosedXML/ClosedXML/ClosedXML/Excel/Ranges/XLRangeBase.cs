using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal abstract class XLRangeBase: IXLRangeBase
    {
        public IXLAddress FirstAddressInSheet { get; protected set; }
        public IXLAddress LastAddressInSheet { get; protected set; }
        internal XLWorksheet Worksheet { get; set; }

        public IXLCell FirstCell()
        {
            return this.Cell(1, 1);
        }
        public IXLCell LastCell()
        {
            return this.Cell(this.RowCount(), this.ColumnCount());
        }

        public IXLCell Cell( IXLAddress cellAddressInRange)
        {
            IXLAddress absoluteAddress = (XLAddress)cellAddressInRange + (XLAddress)this.FirstAddressInSheet - 1;
            if (this.Worksheet.Internals.CellsCollection.ContainsKey(absoluteAddress))
            {
                return this.Worksheet.Internals.CellsCollection[absoluteAddress];
            }
            else
            {
                IXLStyle style = this.Style;
                if (this.Style.ToString() == this.Worksheet.Style.ToString())
                {
                    if (this.Worksheet.Internals.RowsCollection.ContainsKey(absoluteAddress.RowNumber)
                        && this.Worksheet.Internals.RowsCollection[absoluteAddress.RowNumber].Style.ToString() != this.Worksheet.Style.ToString())
                        style = this.Worksheet.Internals.RowsCollection[absoluteAddress.RowNumber].Style;
                    else if (this.Worksheet.Internals.ColumnsCollection.ContainsKey(absoluteAddress.ColumnNumber)
                        && this.Worksheet.Internals.ColumnsCollection[absoluteAddress.ColumnNumber].Style.ToString() != this.Worksheet.Style.ToString())
                        style = this.Worksheet.Internals.ColumnsCollection[absoluteAddress.ColumnNumber].Style;
                }
                var newCell = new XLCell(absoluteAddress, style);
                this.Worksheet.Internals.CellsCollection.Add(absoluteAddress, newCell);
                return newCell;
            }
        }
        public IXLCell Cell( Int32 row, Int32 column)
        {
            return this.Cell(new XLAddress(row, column));
        }
        public IXLCell Cell( Int32 row, String column)
        {
            return this.Cell(new XLAddress(row, column));
        }
        public IXLCell Cell( String cellAddressInRange)
        {
            return this.Cell(new XLAddress(cellAddressInRange));
        }

        public Int32 RowCount()
        {
            return this.LastAddressInSheet.RowNumber - this.FirstAddressInSheet.RowNumber + 1;
        }
        public Int32 ColumnCount()
        {
            return this.LastAddressInSheet.ColumnNumber - this.FirstAddressInSheet.ColumnNumber + 1;
        }

        public IXLRange Range( Int32 firstCellRow, Int32 firstCellColumn, Int32 lastCellRow, Int32 lastCellColumn)
        {
            return this.Range(new XLAddress(firstCellRow, firstCellColumn), new XLAddress(lastCellRow, lastCellColumn));
        }
        public IXLRange Range( String rangeAddress)
        {
            if (rangeAddress.Contains(':'))
            {
                String[] arrRange = rangeAddress.Split(':');
                return this.Range(arrRange[0], arrRange[1]);
            }
            else
            {
                return this.Range(rangeAddress, rangeAddress);
            }
        }
        public IXLRange Range( String firstCellAddress, String lastCellAddress)
        {
            return this.Range(new XLAddress(firstCellAddress), new XLAddress(lastCellAddress));
        }
        public IXLRange Range( IXLAddress firstCellAddress, IXLAddress lastCellAddress)
        {
            var newFirstCellAddress = (XLAddress)firstCellAddress + (XLAddress)this.FirstAddressInSheet - 1;
            var newLastCellAddress = (XLAddress)lastCellAddress + (XLAddress)this.FirstAddressInSheet - 1;
            var xlRangeParameters = new XLRangeParameters(newFirstCellAddress, newLastCellAddress, this.Worksheet, this.Style);
            if (
                   newFirstCellAddress.RowNumber < this.FirstAddressInSheet.RowNumber
                || newFirstCellAddress.RowNumber > this.LastAddressInSheet.RowNumber
                || newLastCellAddress.RowNumber > this.LastAddressInSheet.RowNumber
                || newFirstCellAddress.ColumnNumber < this.FirstAddressInSheet.ColumnNumber
                || newFirstCellAddress.ColumnNumber > this.LastAddressInSheet.ColumnNumber
                || newLastCellAddress.ColumnNumber > this.LastAddressInSheet.ColumnNumber
                )
                throw new ArgumentOutOfRangeException(String.Format("The cells {0} and {1} are outside the range '{2}'.", firstCellAddress.ToString(), lastCellAddress.ToString(), this.ToString()));

            return new XLRange(xlRangeParameters);
        }

        public IXLRanges Ranges( String ranges)
        {
            var retVal = new XLRanges();
            var rangePairs = ranges.Split(',');
            foreach (var pair in rangePairs)
            {
                retVal.Add(this.Range(pair));
            }
            return retVal;
        }
        public IXLRanges Ranges( params String[] ranges)
        {
            var retVal = new XLRanges();
            foreach (var pair in ranges)
            {
                retVal.Add(this.Range(pair));
            }
            return retVal;
        }

        public IEnumerable<IXLCell> Cells()
        {
            foreach (var row in Enumerable.Range(1, this.RowCount()))
            {
                foreach (var column in Enumerable.Range(1, this.ColumnCount()))
                {
                    yield return this.Cell(row, column);
                }
            }
        }
        public IEnumerable<IXLCell> CellsUsed()
        {
            return this.Worksheet.Internals.CellsCollection.Values.AsEnumerable<IXLCell>();
        }

        public void Merge()
        {
            var mergeRange = this.FirstAddressInSheet.ToString() + ":" + this.LastAddressInSheet.ToString();
            if (!this.Worksheet.Internals.MergedCells.Contains(mergeRange))
                this.Worksheet.Internals.MergedCells.Add(mergeRange);
        }
        public void Unmerge()
        {
            this.Worksheet.Internals.MergedCells.Remove(this.FirstAddressInSheet.ToString() + ":" + this.LastAddressInSheet.ToString());
        }

        public abstract IXLStyle Style { get; set; }

        public abstract IEnumerable<IXLStyle> Styles { get; }

        public abstract Boolean UpdatingStyle { get; set; }

        public abstract IXLRange AsRange();
    }
}
