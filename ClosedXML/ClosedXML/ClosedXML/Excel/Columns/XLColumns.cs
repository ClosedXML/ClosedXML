using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLColumns : IXLColumns, IXLStylized
    {
        private XLWorksheet worksheet;
        public XLColumns(XLWorksheet worksheet)
        {
            this.worksheet = worksheet;
            style = new XLStyle(this, XLWorkbook.DefaultStyle);
        }

        List<XLColumn> columns = new List<XLColumn>();
        public IEnumerator<IXLColumn> GetEnumerator()
        {
            var retList = new List<IXLColumn>();
            columns.ForEach(c => retList.Add(c));
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

                if (worksheet != null)
                {
                    worksheet.Style = value;
                }
                else
                {
                    foreach (var column in columns)
                    {
                        column.Style = value;
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
                if (worksheet != null)
                {
                    yield return worksheet.Style;
                }
                else
                {
                    foreach (var col in columns)
                    {
                        foreach (var s in col.Styles)
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

        #endregion

        public Double Width
        {
            set
            {
                columns.ForEach(c => c.Width = value);

                if (worksheet !=null)
                {
                    worksheet.ColumnWidth = value;
                    worksheet.Internals.ColumnsCollection.ForEach(c => c.Value.Width = value);
                }
            }
        }

        public void Delete()
        {
            if (worksheet != null)
            {
                worksheet.Internals.ColumnsCollection.Clear();
                worksheet.Internals.CellsCollection.Clear();
            }
            else
            {
                var toDelete = new Dictionary<IXLWorksheet, List<Int32>>();
                foreach (var c in columns)
                {
                    if (!toDelete.ContainsKey(c.Worksheet))
                        toDelete.Add(c.Worksheet, new List<Int32>());

                    toDelete[c.Worksheet].Add(c.ColumnNumber());
                }

                foreach (var kp in toDelete)
                { 
                    foreach(var c in kp.Value.OrderByDescending(c=>c))
                        kp.Key.Column(c).Delete();
                }
           }
        }

        public void Add(XLColumn column)
        {
            columns.Add(column);
        }

        public IXLColumns AdjustToContents()
        {
            columns.ForEach(c => c.AdjustToContents());
            return this;
        }
        public IXLColumns AdjustToContents(Int32 startRow)
        {
            columns.ForEach(c => c.AdjustToContents(startRow));
            return this;
        }
        public IXLColumns AdjustToContents(Int32 startRow, Int32 endRow)
        {
            columns.ForEach(c => c.AdjustToContents(startRow, endRow));
            return this;
        }

        public IXLColumns AdjustToContents(Double minWidth, Double maxWidth)
        {
            columns.ForEach(c => c.AdjustToContents(minWidth, maxWidth));
            return this;
        }
        public IXLColumns AdjustToContents(Int32 startRow, Double minWidth, Double maxWidth)
        {
            columns.ForEach(c => c.AdjustToContents(startRow, minWidth, maxWidth));
            return this;
        }
        public IXLColumns AdjustToContents(Int32 startRow, Int32 endRow, Double minWidth, Double maxWidth)
        {
            columns.ForEach(c => c.AdjustToContents(startRow, endRow, minWidth, maxWidth));
            return this;
        }

        public void Hide()
        {
            columns.ForEach(c => c.Hide());
        }
        public void Unhide()
        {
            columns.ForEach(c => c.Unhide());
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
            columns.ForEach(c => c.Group(collapse));
        }
        public void Group(Int32 outlineLevel, Boolean collapse)
        {
            columns.ForEach(c => c.Group(outlineLevel, collapse));
        }
        public void Ungroup(Boolean ungroupFromAll)
        {
            columns.ForEach(c => c.Ungroup(ungroupFromAll));
        }
        public void Collapse()
        {
            columns.ForEach(c => c.Collapse());
        }
        public void CollapseOnly()
        {
            columns.ForEach(c => c.Collapsed = true );
        }
        public void Expand()
        {
            columns.ForEach(c => c.Expand());
        }

        public IXLCells Cells()
        {
            var cells = new XLCells(false, false, false);
            foreach (var container in columns)
            {
                cells.Add(container.RangeAddress);
            }
            return (IXLCells)cells;
        }

        public IXLCells CellsUsed()
        {
            var cells = new XLCells(false, true, false);
            foreach (var container in columns)
            {
                cells.Add(container.RangeAddress);
            }
            return (IXLCells)cells;
        }

        public IXLCells CellsUsed(Boolean includeStyles)
        {
            var cells = new XLCells(false, true, includeStyles);
            foreach (var container in columns)
            {
                cells.Add(container.RangeAddress);
            }
            return (IXLCells)cells;
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

        /// <summary>
        /// Adds a vertical page break after this column.
        /// </summary>
        public IXLColumns AddVerticalPageBreaks()
        {
            foreach (var col in columns)
                col.Worksheet.PageSetup.AddVerticalPageBreak(col.ColumnNumber());
            return this;
        }

        public IXLColumns SetDataType(XLCellValues dataType)
        {
            columns.ForEach(c => c.DataType = dataType);
            return this;
        }
    }
}
