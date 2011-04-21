using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLRows : IXLRows, IXLStylized
    {
        private Boolean entireWorksheet;
        private XLWorksheet worksheet;
        public XLRows(XLWorksheet worksheet, Boolean entireWorksheet = false)
        {
            this.worksheet = worksheet;
            this.entireWorksheet = entireWorksheet;
            style = new XLStyle(this, worksheet.Style);
        }

        List<XLRow> rows = new List<XLRow>();

        public IEnumerator<IXLRow> GetEnumerator()
        {
            var retList = new List<IXLRow>();
            rows.ForEach(c => retList.Add(c));
            return retList.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        #region IXLStylized Members

        internal IXLStyle style;
        public IXLStyle Style
        {
            get
            {
                return style;
            }
            set
            {
                style = new XLStyle(this, value);

                if (entireWorksheet)
                {
                    worksheet.Style = value;
                }
                else
                {
                    var maxColumn = 0;
                    if (worksheet.Internals.ColumnsCollection.Count > 0)
                        maxColumn = worksheet.Internals.ColumnsCollection.Keys.Max();

                    foreach (var row in rows)
                    {
                        row.Style = value;
                        foreach (var c in row.Worksheet.Internals.CellsCollection.Values.Where(c => c.Address.RowNumber == row.RangeAddress.FirstAddress.RowNumber))
                        {
                            c.Style = value;
                        }

                        for (var co = 1; co <= maxColumn; co++)
                        {
                            worksheet.Cell(row.RowNumber(), co).Style = value;
                        }
                    }
                }

            }
        }

        public IEnumerable<IXLStyle> Styles
        {
            get
            {
                UpdatingStyle = true;
                yield return style;
                if (entireWorksheet)
                {
                    yield return worksheet.Style;
                }
                else
                {
                    var maxColumn = 0;
                    if (worksheet.Internals.ColumnsCollection.Count > 0)
                        maxColumn = worksheet.Internals.ColumnsCollection.Keys.Max();

                    foreach (var row in rows)
                    {
                        yield return row.Style;
                        foreach (var c in row.Worksheet.Internals.CellsCollection.Values.Where(c => c.Address.RowNumber == row.RangeAddress.FirstAddress.RowNumber))
                        {
                            yield return c.Style;
                        }

                        for (var co = 1; co <= maxColumn; co++)
                        {
                            yield return worksheet.Cell(row.RowNumber(), co).Style;
                        }
                    }
                }
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

        public double Height
        {
            set
            {
                rows.ForEach(c => c.Height = value);
                if (entireWorksheet)
                {
                    worksheet.RowHeight = value;
                    worksheet.Internals.RowsCollection.ForEach(r => r.Value.Height = value);
                }
            }
        }

        public void Delete()
        {
            if (entireWorksheet)
            {
                worksheet.Internals.RowsCollection.Clear();
                worksheet.Internals.CellsCollection.Clear();
            }
            else
            {
                var toDelete = new List<Int32>();
                foreach (var r in rows)
                    toDelete.Add(r.RowNumber());

                foreach (var r in toDelete.OrderByDescending(r => r))
                    worksheet.Row(r).Delete();
            }
        }

        public void Add(XLRow row)
        {
            rows.Add(row);
        }

        public IXLRows AdjustToContents()
        {
            rows.ForEach(r => r.AdjustToContents());
            return this;
        }
        public IXLRows AdjustToContents(Int32 startColumn)
        {
            rows.ForEach(r => r.AdjustToContents(startColumn));
            return this;
        }
        public IXLRows AdjustToContents(Int32 startColumn, Int32 endColumn)
        {
            rows.ForEach(r => r.AdjustToContents(startColumn, endColumn));
            return this;
        }
        public void Hide()
        {
            rows.ForEach(r => r.Hide());
        }
        public void Unhide()
        {
            rows.ForEach(r => r.Unhide());
        }

        public void Group()
        {
            Group(false);
        }
        public void Group(Int32 outlineLevel)
        {
            Group(outlineLevel, false);
        }
        public void Ungroup()
        {
            Ungroup(false);
        }
        public void Group(Boolean collapse)
        {
            rows.ForEach(r => r.Group(collapse));
        }
        public void Group(Int32 outlineLevel, Boolean collapse)
        {
            rows.ForEach(r => r.Group(outlineLevel, collapse));
        }
        public void Ungroup(Boolean ungroupFromAll)
        {
            rows.ForEach(r => r.Ungroup(ungroupFromAll));
        }
        public void Collapse()
        {
            rows.ForEach(r => r.Collapse());
        }
        public void Expand()
        {
            rows.ForEach(r => r.Expand());
        }

        public IXLCells Cells()
        {
            var cells = new XLCells(worksheet, false, false, false);
            foreach (var container in rows)
            {
                cells.Add(container.RangeAddress);
            }
            return (IXLCells)cells;
        }

        public IXLCells CellsUsed()
        {
            var cells = new XLCells(worksheet, false, true, false);
            foreach (var container in rows)
            {
                cells.Add(container.RangeAddress);
            }
            return (IXLCells)cells;
        }

        public IXLCells CellsUsed(Boolean includeStyles)
        {
            var cells = new XLCells(worksheet, false, true, includeStyles);
            foreach (var container in rows)
            {
                cells.Add(container.RangeAddress);
            }
            return (IXLCells)cells;
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

        public IXLRows Replace(String oldValue, String newValue)
        {
            rows.ForEach(r => r.Replace(oldValue, newValue));
            return this;
        }
        public IXLRows Replace(String oldValue, String newValue, XLSearchContents searchContents)
        {
            rows.ForEach(r => r.Replace(oldValue, newValue, searchContents));
            return this;
        }
        public IXLRows Replace(String oldValue, String newValue, XLSearchContents searchContents, Boolean useRegularExpressions)
        {
            rows.ForEach(r => r.Replace(oldValue, newValue, searchContents, useRegularExpressions));
            return this;
        }
    }
}
