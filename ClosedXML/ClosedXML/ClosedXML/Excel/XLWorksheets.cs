using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLWorksheets : IXLWorksheets
    {
        Dictionary<String, IXLWorksheet> worksheets = new Dictionary<String, IXLWorksheet>();

        XLWorkbook workbook;
        public XLWorksheets(XLWorkbook workbook)
        {
            this.workbook = workbook;
        }

        #region IXLWorksheets Members

        public IXLWorksheet Worksheet(string sheetName)
        {
            return worksheets[sheetName];
        }

        public IXLWorksheet Worksheet(int sheetIndex)
        {
            return worksheets.ElementAt(sheetIndex).Value;
        }

        public IXLWorksheet Add(String sheetName)
        {
            var sheet = new XLWorksheet(sheetName, workbook);
            worksheets.Add(sheetName, sheet);
            return sheet;
        }

        public void Delete(string sheetName)
        {
            worksheets.Remove(sheetName);
        }

        public void Delete(Int32 sheetIndex)
        {
            worksheets.Remove(worksheets.ElementAt(sheetIndex).Key);
        }
        
        #endregion

        #region IEnumerable<IXLWorksheet> Members

        public IEnumerator<IXLWorksheet> GetEnumerator()
        {
            return worksheets.Values.GetEnumerator();
        }

        #endregion

        #region IEnumerable Members

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        #endregion


    }
}
