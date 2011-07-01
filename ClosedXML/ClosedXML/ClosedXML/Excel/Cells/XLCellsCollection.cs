using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal class XLCellsCollection
    {
        private const Int32 InitialRowCapacity = 5;
        private const Int32 InitialColumnCapacity = 10;
        //private XLCell[,] _cells;
        private readonly Dictionary<XLSheetPoint, XLCell> _cellsDictionary = new Dictionary<XLSheetPoint, XLCell>();
        private Int32 _columnCapacity = InitialColumnCapacity;
        private Int32 _rowCapacity = InitialRowCapacity;

        public Int32 MaxColumnUsed;
        public Int32 MaxRowUsed;
        public Int32 Count { get; private set; }
        public HashSet<XLSheetPoint> Deleted = new HashSet<XLSheetPoint>();
        public Dictionary<Int32, Int32> RowsUsed = new Dictionary<int, int>();
        public Dictionary<Int32, Int32> ColumnsUsed = new Dictionary<int, int>();

        public XLCellsCollection()
        {
            Clear();
        }

        //private void ResizeIfNecessary(Int32 row, Int32 column)
        //{
        //    if (row >= _rowCapacity || column >= _columnCapacity)
        //    {
        //        if (row >= _rowCapacity)
        //        {
        //            _rowCapacity = (Int32)((Double)_rowCapacity * 2);
                    
        //            if (_rowCapacity < row)
        //                _rowCapacity = (Int32)((Double)row * 1.5);

        //            if (_rowCapacity > ExcelHelper.MaxRowNumber)
        //                _rowCapacity = ExcelHelper.MaxRowNumber;
        //        }

        //        if (column >= _columnCapacity)
        //        {
        //            _columnCapacity = (Int32)((Double)_columnCapacity * 2);

        //            if (_columnCapacity < column)
        //                _columnCapacity = (Int32)((Double)column * 1.5);

        //            if (_columnCapacity > ExcelHelper.MaxColumnNumber)
        //                _columnCapacity = ExcelHelper.MaxColumnNumber;
        //        }
                
        //        _cells = ExcelHelper.ResizeArray(_cells, _rowCapacity + 1, _columnCapacity + 1);
        //    }
        //}

        public void Add(XLSheetPoint sheetPoint, XLCell cell)
        {
            Add(sheetPoint.Row, sheetPoint.Column, cell);
        }
        public void Add(Int32 row, Int32 column, XLCell cell)
        {
            Count++;

            IncrementUsage(RowsUsed, row);
            IncrementUsage(ColumnsUsed, column);

            //ResizeIfNecessary(row, column);
            _cellsDictionary.Add(new XLSheetPoint(row, column),cell );
            //_cells[row, column] = cell;
            if (row > MaxRowUsed) MaxRowUsed = row;
            if (column > MaxColumnUsed) MaxColumnUsed = column;
            var sp = new XLSheetPoint(row, column);
            if (Deleted.Contains(sp))
                Deleted.Remove(sp);
        }

        private static void IncrementUsage(Dictionary<int,int> dictionary, Int32 key)
        {
            if (dictionary.ContainsKey(key))
                dictionary[key]++;
            else
                dictionary.Add(key, 1);
        }

        private static void DecrementUsage(Dictionary<int, int> dictionary, Int32 key)
        {
            Int32 count;
            if (dictionary.TryGetValue(key, out count))
            {
                if (count > 0)
                    dictionary[key]--;
                else
                    dictionary.Remove(key);
            }
        }

        public void Clear()
        {
            Count = 0;
            RowsUsed.Clear();
            ColumnsUsed.Clear();

            //for (int ro = 1; ro <= MaxRowUsed; ro++)
            //{
            //    for (int co = 1; co <= MaxColumnUsed; co++)
            //    {
            //        if (_cells[ro, co] != null)
            //        {
            //            var sp = new XLSheetPoint(ro, co);
            //            if (!Deleted.Contains(sp))
            //                Deleted.Add(sp);
            //        }
            //    }
            //}
            _cellsDictionary.Clear();
            _rowCapacity = InitialRowCapacity;
            _columnCapacity = InitialColumnCapacity;
            MaxRowUsed = 0;
            MaxColumnUsed = 0;
            //_cells = new XLCell[_rowCapacity,_columnCapacity];
        }

        public void Remove(XLSheetPoint sheetPoint)
        {
            Remove(sheetPoint.Row,sheetPoint.Column);
        }
        public void Remove(Int32 row, Int32 column)
        {
            Count--;
            DecrementUsage(RowsUsed, row);
            DecrementUsage(ColumnsUsed, row);
            var sp = new XLSheetPoint(row, column);
            Deleted.Add(sp);
                _cellsDictionary.Remove(sp);
            //_cells[row, column] = null;
        }

        public IEnumerable<XLCell> GetCells(Int32 rowStart, Int32 columnStart,
                                            Int32 rowEnd, Int32 columnEnd)
        {
            int finalRow = rowEnd > MaxRowUsed ? MaxRowUsed : rowEnd;
            int finalColumn = columnEnd > MaxColumnUsed ? MaxColumnUsed : columnEnd;
            for (int ro = rowStart; ro <= finalRow; ro++)
            {
                for (int co = columnStart; co <= finalColumn; co++)
                {
                    XLCell cell;
                    if (_cellsDictionary.TryGetValue(new XLSheetPoint(ro,co), out cell ))
                    {
                        yield return cell;
                    }
                }
            }
        }

        public void RemoveAll(Int32 rowStart, Int32 columnStart,
                                    Int32 rowEnd, Int32 columnEnd)
        {
            int finalRow = rowEnd > MaxRowUsed ? MaxRowUsed : rowEnd;
            int finalColumn = columnEnd > MaxColumnUsed ? MaxColumnUsed : columnEnd;
            for (int ro = rowStart; ro <= finalRow; ro++)
            {
                for (int co = columnStart; co <= finalColumn; co++)
                {
                    var sp = new XLSheetPoint(ro, co);
                    if (_cellsDictionary.ContainsKey(sp))
                    {
                        Remove(sp);
                    }
                }
            }
        }

        public IEnumerable<XLSheetPoint> GetSheetPoints(Int32 rowStart, Int32 columnStart,
                                    Int32 rowEnd, Int32 columnEnd)
        {
            int finalRow = rowEnd > MaxRowUsed ? MaxRowUsed : rowEnd;
            int finalColumn = columnEnd > MaxColumnUsed ? MaxColumnUsed : columnEnd;

            for (int ro = rowStart; ro <= finalRow; ro++)
            {
                for (int co = columnStart; co <= finalColumn; co++)
                {
                    var sp = new XLSheetPoint(ro, co);
                    if (_cellsDictionary.ContainsKey(sp))
                    {
                        yield return sp;
                    }
                }
            }
        }

        public XLCell GetCell(Int32 row, Int32 column)
        {
            if (row > MaxRowUsed || column > MaxColumnUsed)
                return null;
            var sp = new XLSheetPoint(row, column);
            XLCell cell;
            if (_cellsDictionary.TryGetValue(sp, out cell))
                return cell;

            return null;
            //return _cells[row, column];
        }

        public XLCell GetCell(XLSheetPoint sheetPoint)
        {
            XLCell cell;
            if (_cellsDictionary.TryGetValue(sheetPoint, out cell))
                return cell;

            return null;
            //return _cells[sheetPoint.Row, sheetPoint.Column];
        }

        internal void SwapRanges(XLSheetRange sheetRange1, XLSheetRange sheetRange2)
        {
            Int32 rowCount = sheetRange1.LastPoint.Row - sheetRange1.FirstPoint.Row + 1;
            Int32 columnCount = sheetRange1.LastPoint.Column - sheetRange1.FirstPoint.Column + 1;
            for (int row = 0; row < rowCount; row++)
            {
                for (int column = 0; column < columnCount; column++)
                {
                    var sp1 = new XLSheetPoint(sheetRange1.FirstPoint.Row + row, sheetRange1.FirstPoint.Column + column);
                    var sp2 = new XLSheetPoint(sheetRange2.FirstPoint.Row + row, sheetRange2.FirstPoint.Column + column);
                    var cell1 = GetCell(sp1);
                    var cell2 = GetCell(sp2);

                    if (cell1 != null || cell2 != null)
                    {
                        if (cell1 != null)
                        {
                            cell1.Address = new XLAddress(cell1.Worksheet, sp2.Row, sp2.Column, false, false);
                            _cellsDictionary.Remove(sp1);
                            if (cell2 != null)
                                Add(sp1, cell2);
                        }

                        if (cell2 != null)
                        {
                            cell2.Address = new XLAddress(cell2.Worksheet, sp1.Row, sp1.Column, false, false);
                            _cellsDictionary.Remove(sp2);
                            if (cell1 != null)
                                Add(sp2, cell1);
                        }

                    }
                }    
            }
        }

        internal IEnumerable<XLCell> GetCells()
        {
            return GetCells(1, 1, MaxRowUsed, MaxColumnUsed);
        }

        internal IEnumerable<XLCell> GetCells(Func<XLCell, Boolean> predicate)
        {
            for (int ro = 1; ro <= MaxRowUsed; ro++)
            {
                for (int co = 1; co <= MaxColumnUsed; co++)
                {
                    var cell = GetCell(ro, co);
                    if (cell != null && predicate(cell))
                    {
                        yield return cell;
                    }
                }
            }
        }

        public Boolean Contains(Int32 row, Int32 column)
        {
            return _cellsDictionary.ContainsKey(new XLSheetPoint(row, column));
        }

        public Int32 MinRowInColumn(Int32 column)
        {
            for (int row = 1; row <= MaxRowUsed; row++)
            {
                if (GetCell(row, column) != null)
                    return row;
            }

            return 0;
        }

        public Int32 MaxRowInColumn(Int32 column)
        {
            for (int row = MaxRowUsed; row >= 1; row--)
            {
                if (GetCell(row, column) != null)
                    return row;
            }

            return 0;
        }

        public Int32 MinColumnInRow(Int32 row)
        {
            for (int column = 1; column <= MaxColumnUsed; column++)
            {
                if (GetCell(row, column) != null)
                    return column;
            }

            return 0;
        }

        public Int32 MaxColumnInRow(Int32 row)
        {
            for (int column = MaxColumnUsed; column >= 1; column--)
            {
                if (GetCell(row, column) != null)
                    return column;
            }

            return 0;
        }

        public IEnumerable<XLCell> GetCellsInColumn(Int32 column)
        {
            return GetCells(1, column, MaxRowUsed, column);
        }

        public IEnumerable<XLCell> GetCellsInRow(Int32 row)
        {
            return GetCells(row, 1, row, MaxColumnUsed);
        }
    }
}