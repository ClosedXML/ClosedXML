using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLCells : IXLCells, IXLStylized
    {
        private XLWorksheet worksheet;
        private Boolean usedCellsOnly;
        private Boolean includeStyles;
        public XLCells(XLWorksheet worksheet, Boolean entireWorksheet, Boolean usedCellsOnly, Boolean includeStyles)
        {
            this.worksheet = worksheet;
            this.style = new XLStyle(this, worksheet.Style);
            this.usedCellsOnly = usedCellsOnly;
            this.includeStyles = includeStyles;
        }

        private List<IXLRangeAddress> rangeAddresses = new List<IXLRangeAddress>();
        public IEnumerator<IXLCell> GetEnumerator()
        {
            HashSet<IXLAddress> usedCells;
            Boolean multipleRanges = rangeAddresses.Count > 1;

            if (multipleRanges)
                usedCells = new HashSet<IXLAddress>();
            else
                usedCells = null;


            if (usedCellsOnly)
            {
                var cells = from c in worksheet.Internals.CellsCollection
                            where (   !StringExtensions.IsNullOrWhiteSpace(c.Value.InnerText)
                                  || (includeStyles && !c.Value.Style.Equals(worksheet.Style)))
                                  && rangeAddresses.FirstOrDefault(r=>
                                      r.FirstAddress.RowNumber <= c.Key.RowNumber
                                      && r.FirstAddress.ColumnNumber <= c.Key.ColumnNumber
                                      && r.LastAddress.RowNumber >= c.Key.RowNumber
                                      && r.LastAddress.ColumnNumber >= c.Key.ColumnNumber
                                      ) != null
                            select (IXLCell)c.Value;
                foreach (var cell in cells)
                {
                    yield return cell;
                }
            }
            else
            {
                foreach (var range in rangeAddresses)
                {
                    Int32 firstRo = range.FirstAddress.RowNumber;
                    Int32 lastRo = range.LastAddress.RowNumber;
                    Int32 firstCo = range.FirstAddress.ColumnNumber;
                    Int32 lastCo = range.LastAddress.ColumnNumber;

                    for (Int32 ro = firstRo; ro <= lastRo; ro++)
                    {
                        for (Int32 co = firstCo; co <= lastCo; co++)
                        {
                            var cell = worksheet.Cell(ro, co);
                            if (multipleRanges)
                            {
                                if (!usedCells.Contains(cell.Address))
                                {
                                    usedCells.Add(cell.Address);
                                    yield return cell;
                                }
                            }
                            else
                            {
                                yield return cell;
                            }
                        }
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
                var retVal = new XLRanges(worksheet.Internals.Workbook, this.Style);
                this.ForEach(c => retVal.Add(c.AsRange()));
                return retVal;
            }
        }
    }
}
