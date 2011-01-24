using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLWorksheets : IXLWorksheets
    {
        Dictionary<String, XLWorksheet> worksheets = new Dictionary<String, XLWorksheet>();
        public HashSet<String> Deleted = new HashSet<String>();
        XLWorkbook workbook;
        public XLWorksheets(XLWorkbook workbook)
        {
            this.workbook = workbook;
        }

        #region IXLWorksheets Members

        public IXLWorksheet Worksheet(String sheetName)
        {
            return worksheets[sheetName];
        }

        public IXLWorksheet Worksheet(Int32 position)
        {
            var wsCount = worksheets.Values.Where(w => w.Position == position).Count();
            if (wsCount == 0)
                throw new Exception("There isn't a worksheet associated with that position.");

            if (wsCount > 1)
                throw new Exception("Can't retrieve a worksheet because there are multiple worksheets associated with that position.");

            return worksheets.Values.Where(w => w.Position == position).Single();
        }

        public IXLWorksheet Add(String sheetName)
        {
            var sheet = new XLWorksheet(sheetName, workbook);
            worksheets.Add(sheetName, sheet);
            sheet.position = worksheets.Count;
            return sheet;
        }

        public IXLWorksheet Add(String sheetName, Int32 position)
        {
            var ws = Add(sheetName);
            ws.Position = position;
            return ws;
        }

        public void Delete(String sheetName)
        {
            Delete(worksheets[sheetName].Position);
        }

        public void Delete(Int32 position)
        {
            var wsCount = worksheets.Values.Where(w => w.Position == position).Count();
            if (wsCount == 0)
                throw new Exception("There isn't a worksheet associated with that index.");

            if (wsCount > 1)
                throw new Exception("Can't delete the worksheet because there are multiple worksheets associated with that index.");

            var ws = (XLWorksheet)worksheets.Values.Where(w => w.Position == position).Single();
            if (!StringExtensions.IsNullOrWhiteSpace(ws.RelId) && !Deleted.Contains(ws.RelId))
                Deleted.Add(ws.RelId);

            worksheets.RemoveAll(w => w.Position == position);
            worksheets.Values.Where(w => w.Position > position).ForEach(w => ((XLWorksheet)w).position -= 1);
        }
        
        #endregion

        #region IEnumerable<IXLWorksheet> Members

        public IEnumerator<IXLWorksheet> GetEnumerator()
        {
            foreach (var w in worksheets.Values)
            { 
                yield return (IXLWorksheet)w;
            }
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
