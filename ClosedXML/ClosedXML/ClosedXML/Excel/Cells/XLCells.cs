using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLCells: IXLCells
    {
        private Boolean entireWorksheet;
        private XLWorksheet worksheet;
        public XLCells(XLWorksheet worksheet, Boolean entireWorksheet = false)
        {
            this.worksheet = worksheet;
            this.entireWorksheet = entireWorksheet;
            Style = worksheet.Style;
        }

        private List<IXLCell> cells = new List<IXLCell>();
        public IEnumerator<IXLCell> GetEnumerator()
        {
            var retList = new List<IXLCell>();
            cells.ForEach(c => retList.Add(c));
            return retList.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
        public void Add(XLCell cell)
        {
            cells.Add(cell);
        }
        public void AddRange(IEnumerable<IXLCell> cellsToAdd)
        {
            cells.AddRange(cellsToAdd);
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

                if (entireWorksheet)
                {
                    worksheet.Style = value;
                }
                else
                {
                    var maxRow = 0;
                    if (worksheet.Internals.RowsCollection.Count > 0)
                        maxRow = worksheet.Internals.RowsCollection.Keys.Max();
                    foreach (var c in cells)
                    {
                        c.Style = value;
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
                    var maxRow = 0;
                    if (worksheet.Internals.RowsCollection.Count > 0)
                        maxRow = worksheet.Internals.RowsCollection.Keys.Max();
                    foreach (var c in cells)
                    {
                        yield return c.Style;
                    }
                }
                UpdatingStyle = false;
            }
        }

        public Boolean UpdatingStyle { get; set; }

        #endregion
    }
}
