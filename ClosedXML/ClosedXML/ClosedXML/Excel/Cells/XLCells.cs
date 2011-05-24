using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLCells : IXLCells, IXLStylized
    {
        private Boolean usedCellsOnly;
        private Boolean includeStyles;
        public XLCells(Boolean entireWorksheet, Boolean usedCellsOnly, Boolean includeStyles)
        {
            this.style = new XLStyle(this, XLWorkbook.DefaultStyle);
            this.usedCellsOnly = usedCellsOnly;
            this.includeStyles = includeStyles;
        }

        private List<IXLRangeAddress> rangeAddresses = new List<IXLRangeAddress>();
        private struct MinMax { public Int32 MinRow; public Int32 MaxRow; public Int32 MinColumn; public Int32 MaxColumn; }
        public IEnumerator<IXLCell> GetEnumerator()
        {
            var cellsInRanges = new Dictionary<IXLWorksheet, HashSet<IXLAddress>>();
            foreach (var range in rangeAddresses)
            {
                HashSet<IXLAddress> hash;
                if (cellsInRanges.ContainsKey(range.Worksheet))
                {
                    hash = cellsInRanges[range.Worksheet];
                }
                else
                {
                    hash = new HashSet<IXLAddress>();
                    cellsInRanges.Add(range.Worksheet, hash);
                }

                if (usedCellsOnly)
                {
                    var addressList = (range.Worksheet as XLWorksheet).Internals.CellsCollection.Keys.Where(a =>
                            a.RowNumber >= range.FirstAddress.RowNumber
                            && a.RowNumber <= range.LastAddress.RowNumber
                            && a.ColumnNumber >= range.FirstAddress.ColumnNumber
                            && a.ColumnNumber <= range.LastAddress.ColumnNumber).Select(a => a);

                    foreach (var a in addressList)
                    {
                        if (!hash.Contains(a))
                            hash.Add(a);
                    }

                }
                else
                {
                    var mm = new MinMax();
                    mm.MinRow = range.FirstAddress.RowNumber;
                    mm.MaxRow = range.LastAddress.RowNumber;
                    mm.MinColumn = range.FirstAddress.ColumnNumber;
                    mm.MaxColumn = range.LastAddress.ColumnNumber;
                    if (mm.MaxRow > 0 && mm.MaxColumn > 0)
                    {
                        for (Int32 ro = mm.MinRow; ro <= mm.MaxRow; ro++)
                        {
                            for (Int32 co = mm.MinColumn; co <= mm.MaxColumn; co++)
                            {
                                var address = new XLAddress(range.Worksheet, ro, co, false, false);
                                if (!hash.Contains(address))
                                    hash.Add(address);
                            }
                        }
                    }
                }
            }

            if (usedCellsOnly)
            {
                foreach (var cir in cellsInRanges)
                {
                    var cellsCollection = (cir.Key as XLWorksheet).Internals.CellsCollection;
                    foreach (var a in cir.Value)
                    {
                        if (cellsCollection.ContainsKey(a))
                        { 
                            var cell = cellsCollection[a];
                            if (!StringExtensions.IsNullOrWhiteSpace((cell as XLCell).InnerText)
                                || (includeStyles && !cell.Style.Equals(cir.Key.Style)))
                            {
                                yield return cell;
                            }
                        }
                    }

                    //foreach (var cell in (cir.Key as XLWorksheet).Internals.CellsCollection
                    //    .Where(kp => cir.Value.Contains(kp.Key)
                    //        && (!StringExtensions.IsNullOrWhiteSpace((kp.Value as XLCell).InnerText)
                    //            || (includeStyles && !kp.Value.Style.Equals(cir.Key.Style))))
                    //    .Select(kp => kp.Value))
                    //{
                    //    yield return cell;
                    //}
                }
            }
            else
            {
                foreach (var cir in cellsInRanges)
                {
                    foreach (var address in cir.Value)
                    {
                        yield return cir.Key.Cell(address);
                    }
                }
            }
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
        public void Add(IXLRangeAddress rangeAddress)
        {
            rangeAddresses.Add(rangeAddress);
        }

        public void Add(IXLCell cell)
        {
            rangeAddresses.Add(new XLRangeAddress(cell.Address, cell.Address));
        }

        #region IXLStylized Members

        private IXLStyle style;
        public IXLStyle Style
        {
            get
            {
                return style;
            }
            set
            {
                style = new XLStyle(this, value);
                this.ForEach(c => c.Style = style);
            }
        }

        public IEnumerable<IXLStyle> Styles
        {
            get
            {
                UpdatingStyle = true;
                yield return style;
                foreach (var c in this)
                    yield return c.Style;
                UpdatingStyle = false;
            }
        }

        public Boolean UpdatingStyle { get; set; }

        public IXLStyle InnerStyle
        {
            get { return style; }
            set { style = new XLStyle(this, value); }
        }

        #endregion

        public Object Value 
        {
            set
            {
                this.ForEach(c => c.Value = value);
            }
        }

        public XLCellValues DataType
        {
            set
            {
                this.ForEach(c => c.DataType = value);
            }
        }

        public void Clear()
        {
            this.ForEach(c => c.Clear());
        }

        public void ClearStyles()
        {
            this.ForEach(c => c.ClearStyles());
        }

        public String FormulaA1
        {
            set
            {
                this.ForEach(c => c.FormulaA1 = value);
            }
        }

        public String FormulaR1C1
        {
            set
            {
                this.ForEach(c => c.FormulaR1C1 = value);
            }
        }

        public IXLRanges RangesUsed
        {
            get
            {
                var retVal = new XLRanges();
                this.ForEach(c => retVal.Add(c.AsRange()));
                return retVal;
            }
        }

  
    }
}
