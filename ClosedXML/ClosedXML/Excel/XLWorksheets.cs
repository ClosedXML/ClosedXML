using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public class XLWorksheets : IEnumerable<XLWorksheet>
    {
        private Dictionary<String, XLWorksheet> worksheets = new Dictionary<String, XLWorksheet>();

        private XLWorkbook workbook;

        public XLWorksheets(XLWorkbook workbook)
        {
            this.workbook = workbook;
        }

        public XLWorksheet this[String sheetName]
        {
            get
            {
                return worksheets[sheetName];
            }
        }

        public XLWorksheet Add(String name)
        {
            XLWorksheet worksheet = new XLWorksheet(workbook,  name, new XLCells(workbook));
            worksheets.Add(name, worksheet);
            return worksheet;
        }

        public UInt32 Count
        {
            get
            {
                return (UInt32)worksheets.Count;
            }
        }

        private Int32 nextWorksheetId = 1;
        private Int32 GetNextWorksheetId()
        {
            return nextWorksheetId++;
        }

        #region IEnumerable<XLWorksheet> Members

        public IEnumerator<XLWorksheet> GetEnumerator()
        {
            return worksheets.Values.AsEnumerable().GetEnumerator();
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
