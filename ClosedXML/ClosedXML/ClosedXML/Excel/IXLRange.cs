using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel.Style;

namespace ClosedXML.Excel
{
    public interface IXLRange: IXLStylized
    {
        Dictionary<IXLAddress, IXLCell> CellsCollection { get; }
        List<String> MergedCells { get; }

        IXLAddress FirstCellAddress { get; }
        IXLAddress LastCellAddress { get; }
        IXLRange Row(Int32 row);
        IXLRange Column(Int32 column);
        IXLRange Column(String column);
    }

    public static class XLRangeMethods
    {
        public static IXLCell FirstCell(this IXLRange range)
        {
            return range.Cell(1, 1);
        }
        public static IXLCell LastCell(this IXLRange range)
        {
            return range.Cell(range.RowCount(), range.ColumnCount());
        }

        public static IXLCell Cell(this IXLRange range, IXLAddress cellAddressInRange)
        {
            IXLAddress absoluteAddress = (XLAddress)cellAddressInRange + (XLAddress)range.FirstCellAddress - 1;
            if (range.CellsCollection.ContainsKey(absoluteAddress))
            {
                return range.CellsCollection[absoluteAddress];
            }
            else
            {
                var newCell = new XLCell(absoluteAddress, range.Style);
                range.CellsCollection.Add(absoluteAddress, newCell);
                return newCell;
            }
        }
        public static IXLCell Cell(this IXLRange range, Int32 row, Int32 column)
        {
            return range.Cell(new XLAddress(row, column));
        }
        public static IXLCell Cell(this IXLRange range, Int32 row, String column)
        {
            return range.Cell(new XLAddress(row, column));
        }
        public static IXLCell Cell(this IXLRange range, String cellAddressInRange)
        {
            return range.Cell(new XLAddress(cellAddressInRange));
        }

        public static Int32 RowCount(this IXLRange range)
        {
            return range.LastCellAddress.Row - range.FirstCellAddress.Row + 1;
        }
        public static Int32 ColumnCount(this IXLRange range)
        {
            return range.LastCellAddress.Column - range.FirstCellAddress.Column + 1;
        }

        public static XLRange Range(this IXLRange range, String rangeAddress)
        {
            String[] arrRange = rangeAddress.Split(':');
            return range.Range(arrRange[0], arrRange[1]);
        }
        public static XLRange Range(this IXLRange range, String firstCellAddress, String lastCellAddress)
        {
            return range.Range(new XLAddress(firstCellAddress), new XLAddress(lastCellAddress));
        }
        public static XLRange Range(this IXLRange range, IXLAddress firstCellAddress, IXLAddress lastCellAddress)
        {
            return new XLRange(
                new XLRangeParameters()
                {
                    FirstCellAddress = (XLAddress)firstCellAddress + (XLAddress)range.FirstCellAddress - 1,
                    LastCellAddress = (XLAddress)lastCellAddress + (XLAddress)range.FirstCellAddress - 1,
                    CellsCollection = range.CellsCollection,
                    MergedCells = range.MergedCells
                }
                );
        }
        public static XLRange Range(this IXLRange range, IXLCell firstCell, IXLCell lastCell)
        {
            return range.Range(firstCell.Address, lastCell.Address);
        }


        public static IEnumerable<IXLCell> Cells(this IXLRange range)
        {
            foreach(var row in Enumerable.Range(1, range.RowCount())) 
            {
                foreach(var column in Enumerable.Range(1, range.ColumnCount()))
                {
                    yield return range.Cell(row, column);
                }
            }
        }

       
    }
}

