using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    using System.Collections;

    internal class XLRows : IXLRows, IXLStylized
    {
        private readonly List<XLRow> rows = new List<XLRow>();
        private readonly XLWorksheet worksheet;
        internal IXLStyle style;

        public XLRows(XLWorksheet worksheet)
        {
            this.worksheet = worksheet;
            style = new XLStyle(this, XLWorkbook.DefaultStyle);
        }

        #region IXLRows Members

        public IEnumerator<IXLRow> GetEnumerator()
        {
            var retList = new List<IXLRow>();
            rows.ForEach(retList.Add);
            return retList.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public IXLStyle Style
        {
            get { return style; }
            set
            {
                style = new XLStyle(this, value);

                if (worksheet != null)
                    worksheet.Style = value;
                else
                {
                    foreach (XLRow row in rows)
                        row.Style = value;
                }
            }
        }

        public double Height
        {
            set
            {
                rows.ForEach(c => c.Height = value);
                if (worksheet != null)
                {
                    worksheet.RowHeight = value;
                    worksheet.Internals.RowsCollection.ForEach(r => r.Value.Height = value);
                }
            }
        }

        public void Delete()
        {
            if (worksheet != null)
            {
                worksheet.Internals.RowsCollection.Clear();
                worksheet.Internals.CellsCollection.Clear();
            }
            else
            {
                var toDelete = new Dictionary<IXLWorksheet, List<Int32>>();
                foreach (XLRow r in rows)
                {
                    if (!toDelete.ContainsKey(r.Worksheet))
                        toDelete.Add(r.Worksheet, new List<Int32>());

                    toDelete[r.Worksheet].Add(r.RowNumber());
                }

                foreach (KeyValuePair<IXLWorksheet, List<int>> kp in toDelete)
                {
                    foreach (int r in kp.Value.OrderByDescending(r => r))
                        kp.Key.Row(r).Delete();
                }
            }
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

        public IXLRows AdjustToContents(Double minHeight, Double maxHeight)
        {
            rows.ForEach(r => r.AdjustToContents(minHeight, maxHeight));
            return this;
        }

        public IXLRows AdjustToContents(Int32 startColumn, Double minHeight, Double maxHeight)
        {
            rows.ForEach(r => r.AdjustToContents(startColumn, minHeight, maxHeight));
            return this;
        }

        public IXLRows AdjustToContents(Int32 startColumn, Int32 endColumn, Double minHeight, Double maxHeight)
        {
            rows.ForEach(r => r.AdjustToContents(startColumn, endColumn, minHeight, maxHeight));
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
            var cells = new XLCells(false, false);
            foreach (XLRow container in rows)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLCells CellsUsed()
        {
            var cells = new XLCells(true, false);
            foreach (XLRow container in rows)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLCells CellsUsed(Boolean includeStyles)
        {
            var cells = new XLCells(true, includeStyles);
            foreach (XLRow container in rows)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLRows AddHorizontalPageBreaks()
        {
            foreach (XLRow row in rows)
                row.Worksheet.PageSetup.AddHorizontalPageBreak(row.RowNumber());
            return this;
        }

        public IXLRows SetDataType(XLCellValues dataType)
        {
            rows.ForEach(c => c.DataType = dataType);
            return this;
        }

        #endregion

        #region IXLStylized Members

        public IEnumerable<IXLStyle> Styles
        {
            get
            {
                UpdatingStyle = true;
                yield return style;
                if (worksheet != null)
                    yield return worksheet.Style;
                else
                {
                    foreach (XLRow row in rows)
                    {
                        foreach (IXLStyle s in row.Styles)
                            yield return s;
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

        public IXLRanges RangesUsed
        {
            get
            {
                var retVal = new XLRanges();
                this.ForEach(c => retVal.Add(c.AsRange()));
                return retVal;
            }
        }

        #endregion

        public void Add(XLRow row)
        {
            rows.Add(row);
        }
    }
}