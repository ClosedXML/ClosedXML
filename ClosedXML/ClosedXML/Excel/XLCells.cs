using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public class XLCells: IEnumerable<XLCell>
    {
        private Dictionary<XLCellAddress, XLCell> cells = new Dictionary<XLCellAddress, XLCell>();
        private XLWorkbook workbook;
        public XLCells(XLWorkbook workbook)
        {
            this.workbook = workbook;
        }

        public XLCell this[XLCellAddress cellAddress]
        {
            get
            {
                Add(cellAddress);
                return cells[cellAddress];
            }
        }

        public void Add(XLCellAddress cellAddress)
        {
            if (!cells.ContainsKey(cellAddress))
                cells.Add(cellAddress, new XLCell(workbook, cellAddress));
        }

        #region IEnumerable<XLCell> Members

        public IEnumerator<XLCell> GetEnumerator()
        {
            return cells.Values.AsEnumerable().GetEnumerator();
        }

        #endregion

        #region IEnumerable Members

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        #endregion
    }
}
