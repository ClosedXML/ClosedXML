using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public class XLWorksheets: IXLWorksheets
    {
        Dictionary<String, IXLWorksheet> worksheets = new Dictionary<String, IXLWorksheet>();

        #region IXLWorksheets Members

        public IXLWorksheet GetWorksheet(string sheetName)
        {
            return worksheets[sheetName];
        }

        public IXLWorksheet GetWorksheet(int sheetIndex)
        {
            return worksheets.ElementAt(sheetIndex).Value;
        }

        public IXLWorksheet Add(String sheetName)
        {
            var sheet = new XLWorksheet(sheetName);
            worksheets.Add(sheetName, sheet);
            return sheet;
        }

        public void Delete(string sheetName)
        {
            worksheets.Remove(sheetName);
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
