using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public enum CellContent { All, WithValues }
    public class XLRange
    {
        private XLCell firstCell { get; set; }
        public XLCell FirstCell()
        {
            return FirstCell(CellContent.All);
        }
        public XLCell FirstCell(CellContent cellContent)
        {
            if (cellContent == CellContent.WithValues)
            {
                var cellsWithValues = Cells(cellContent);
                var minAddress = cellsWithValues.Min(c => c.CellAddress);
                return new XLCell(this.workbook, minAddress);
            }
            else
            {
                return firstCell;
            }
        }

        private XLCell lastCell { get; set; }
        public XLCell LastCell()
        {
            return LastCell(CellContent.All);
        }
        public XLCell LastCell(CellContent cellContent)
        {
            if (cellContent == CellContent.WithValues)
            {
                var cellsWithValues = Cells(cellContent);
                var maxAddress = cellsWithValues.Max(c => c.CellAddress);
                return new XLCell(this.workbook, maxAddress);
            }
            else
            {
                return lastCell;
            }
        }

        private XLCells cellData;
        public XLRange ParentRange { get; private set; }
        protected XLWorkbook workbook;

        internal XLRange(XLCell firstCell, XLCell lastCell, XLCells cellData, XLRange parentRange)
        {
            this.cellData = cellData;
            this.firstCell = firstCell;
            this.lastCell = lastCell;
            this.ParentRange = parentRange;
            this.workbook = parentRange == null ? null : parentRange.workbook;
        }

        public XLCell Cell(UInt32 row, String column)
        {
            return Cell(column + row.ToString());
        }

        public XLCell Cell(UInt32 row, UInt32 column)
        {
            XLCellAddress cellAddress = new XLCellAddress(row-1, column-1) + firstCell.CellAddress;
            return cellData[cellAddress];
        }

        public XLCell Cell(String cellAddressString)
        {
            XLCellAddress cellAddress = new XLCellAddress(cellAddressString);
            return cellData[cellAddress];
        }

        public XLRange Range(XLCell firstCell, XLCell lastCell)
        {
            return new XLRange(firstCell, lastCell, cellData, this);
        }

        public XLRange Range(String range)
        {
            String[] ranges = range.Split(':');
            XLCell firstCell = new XLCell(workbook, ranges[0] );
            XLCell lastCell = new XLCell(workbook, ranges[1] );
            return Range(firstCell, lastCell);
        }

        public Boolean HasData { get { return cellData.Any(); } }


        public UInt32 CellCount()
        {
            return CellCount(CellContent.All);
        }
        public UInt32 CellCount(CellContent cellContent)
        {
            return (UInt32)Cells(cellContent).Count;
        }

        public UInt32 Row { get { return firstCell.Row; } }
        public List<XLRange> Rows()
        {
            return Rows(CellContent.All);
        }
        public List<XLRange> Rows(CellContent cellContent)
        {
            if (cellContent == CellContent.WithValues)
            {
                var cellsWithValues = Cells(CellContent.WithValues);
                var distinct = cellsWithValues.Select(c => c.Row).Distinct();
                var rows = from d in distinct
                           select new XLRange(
                                  new XLCell(workbook, new XLCellAddress(d, 1))
                                , new XLCell(workbook, new XLCellAddress(d, XLWorksheet.MaxNumberOfColumns))
                                , cellData, this);
                return rows.ToList();
            }
            else
            {
                var distinct = Cells().Select(c => c.Row).Distinct();
                var rows = from d in distinct
                           select new XLRange(
                                  new XLCell(workbook, new XLCellAddress(d, 1))
                                , new XLCell(workbook, new XLCellAddress(d, XLWorksheet.MaxNumberOfColumns))
                                , cellData, this);
                return rows.ToList();
            }
        }
        public UInt32 RowCount()
        {
            return RowCount(CellContent.All);
        }
        public UInt32 RowCount(CellContent cellContent)
        {
            if (cellContent == CellContent.WithValues)
            {
                var cellsWithValues = Cells(CellContent.WithValues);
                var distinct = cellsWithValues.Select(c => c.Row).Distinct();
                return (UInt32)(distinct.Count());
            }
            else
            {
                return (lastCell.CellAddress - firstCell.CellAddress).Row + 1;
            }
        }

        public UInt32 Column { get { return firstCell.Column; } }
        public UInt32 ColumnCount()
        {
            return ColumnCount(CellContent.All);
        }
        public UInt32 ColumnCount(CellContent cellContent)
        {
            if (cellContent == CellContent.WithValues)
            {
                var cellsWithValues = Cells(CellContent.WithValues);
                var distinct = cellsWithValues.Select(c => c.Column).Distinct();
                return (UInt32)(distinct.Count());
            }
            else
            {
                return (lastCell.CellAddress - firstCell.CellAddress).Column + 1;
            }
        }

        public virtual List<XLCell> Cells()
        {
            return Cells(CellContent.All);
        }

        public virtual List<XLCell> Cells(CellContent cellContent)
        {
            if (cellContent == CellContent.WithValues)
            {
                return cellData
                        .Where(c => c.HasValue && c.CellAddress >= this.firstCell.CellAddress && c.CellAddress <= lastCell.CellAddress)
                        .OrderBy(x => x.CellAddress)
                        .ToList();
            }
            else
            {
                List<XLCell> retVal = new List<XLCell>();
                for (UInt32 row = firstCell.Row; row <= lastCell.Row; row++)
                {
                    for (UInt32 column = firstCell.Column; column <= lastCell.Column; column++)
                    {
                        retVal.Add(Cell(row, column));
                    }
                }
                return retVal;
            }
        }
    }
}
