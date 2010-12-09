using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLRows: IXLRows
    {
        private Boolean entireWorksheet;
        private XLWorksheet worksheet;
        public XLRows(XLWorksheet worksheet, Boolean entireWorksheet = false)
        {
            this.worksheet = worksheet;
            this.entireWorksheet = entireWorksheet;
            Style = worksheet.Style;
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
                worksheet.Internals.ColumnsCollection.Clear();
                worksheet.Internals.CellsCollection.Clear();
            }
            else
            {
                rows.ForEach(r => r.Delete());
            }
        }

        public void Add(XLRow row)
        {
            rows.Add(row);
        }

        public void AdjustToContents()
        {
            rows.ForEach(r => r.AdjustToContents());
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
    }
}
