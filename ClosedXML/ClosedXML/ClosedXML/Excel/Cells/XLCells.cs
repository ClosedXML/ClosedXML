using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLCells : IXLCells, IXLStylized, IEnumerable<XLCell>
    {
        #region Fields
        private readonly bool m_usedCellsOnly;
        private readonly bool m_includeStyles;
        private readonly List<XLRangeAddress> m_rangeAddresses = new List<XLRangeAddress>();
        private IXLStyle m_style;
        #endregion
        #region Constructor
        public XLCells(bool entireWorksheet, bool usedCellsOnly, bool includeStyles)
        {
            m_style = new XLStyle(this, XLWorkbook.DefaultStyle);
            m_usedCellsOnly = usedCellsOnly;
            m_includeStyles = includeStyles;
        }
        #endregion
        public IEnumerator<XLCell> GetEnumerator()
        {
            var cellsInRanges = new Dictionary<XLWorksheet, HashSet<IXLAddress>>();
            foreach (var range in m_rangeAddresses)
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

                if (m_usedCellsOnly)
                {
                    var tmpRange = range;
                    var addressList = range.Worksheet.Internals.CellsCollection.Keys
                            .Where(a => a.RowNumber >= tmpRange.FirstAddress.RowNumber &&
                                        a.RowNumber <= tmpRange.LastAddress.RowNumber &&
                                        a.ColumnNumber >= tmpRange.FirstAddress.ColumnNumber &&
                                        a.ColumnNumber <= tmpRange.LastAddress.ColumnNumber);

                    foreach (var a in addressList)
                    {
                        if (!hash.Contains(a))
                        {
                            hash.Add(a);
                        }
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
                                {
                                    hash.Add(address);
                                }
                            }
                        }
                    }
                }
            }

            if (m_usedCellsOnly)
            {
                foreach (var cir in cellsInRanges)
                {
                    var cellsCollection = cir.Key.Internals.CellsCollection;
                    foreach (var a in cir.Value)
                    {
                        if (cellsCollection.ContainsKey(a))
                        {
                            var cell = cellsCollection[a];
                            if (!StringExtensions.IsNullOrWhiteSpace((cell).InnerText)
                                || (m_includeStyles && !cell.Style.Equals(cir.Key.Style)))
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
        IEnumerator<IXLCell> IEnumerable<IXLCell>.GetEnumerator()
        {
            foreach (var cell in this)
            {
                yield return cell;
            }
        }
        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
        public void Add(XLRangeAddress rangeAddress)
        {
            m_rangeAddresses.Add(rangeAddress);
        }

        public void Add(XLCell cell)
        {
            m_rangeAddresses.Add(new XLRangeAddress(cell.Address, cell.Address));
        }
        #region IXLStylized Members
        public IXLStyle Style
        {
            get { return m_style; }
            set
            {
                m_style = new XLStyle(this, value);
                this.ForEach<XLCell>(c => c.Style = m_style);
            }
        }

        public IEnumerable<IXLStyle> Styles
        {
            get
            {
                UpdatingStyle = true;
                yield return m_style;
                foreach (var c in this)
                {
                    yield return c.Style;
                }
                UpdatingStyle = false;
            }
        }

        public Boolean UpdatingStyle { get; set; }

        public IXLStyle InnerStyle
        {
            get { return m_style; }
            set { m_style = new XLStyle(this, value); }
        }
        #endregion
        public Object Value
        {
            set { this.ForEach<XLCell>(c => c.Value = value); }
        }

        public IXLCells SetDataType(XLCellValues dataType)
        {
            this.ForEach<XLCell>(c => c.DataType = dataType);
            return this;
        }

        public XLCellValues DataType
        {
            set { this.ForEach<XLCell>(c => c.DataType = value); }
        }

        public void Clear()
        {
            this.ForEach<XLCell>(c => c.Clear());
        }

        public void ClearStyles()
        {
            this.ForEach<XLCell>(c => c.ClearStyles());
        }

        public String FormulaA1
        {
            set { this.ForEach<XLCell>(c => c.FormulaA1 = value); }
        }

        public String FormulaR1C1
        {
            set { this.ForEach<XLCell>(c => c.FormulaR1C1 = value); }
        }

        public IXLRanges RangesUsed
        {
            get
            {
                var retVal = new XLRanges();
                this.ForEach<XLCell>(c => retVal.Add(c.AsRange()));
                return retVal;
            }
        }
        //--
        #region  Nested type: MinMax
        private struct MinMax
        {
            public Int32 MinRow;
            public Int32 MaxRow;
            public Int32 MinColumn;
            public Int32 MaxColumn;
        }
        #endregion
    }
}