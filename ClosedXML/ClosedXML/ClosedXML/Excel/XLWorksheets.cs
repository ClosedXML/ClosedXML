using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLWorksheets : IXLWorksheets, IEnumerable<XLWorksheet>
    {
        #region Constructor

        private readonly XLWorkbook _workbook;
        private readonly Dictionary<String, XLWorksheet> _worksheets = new Dictionary<String, XLWorksheet>();

        #endregion

        public HashSet<String> Deleted = new HashSet<String>();

        #region Constructor

        public XLWorksheets(XLWorkbook workbook)
        {
            _workbook = workbook;
        }

        #endregion

        #region IEnumerable<XLWorksheet> Members

        public IEnumerator<XLWorksheet> GetEnumerator()
        {
            return ((IEnumerable<XLWorksheet>)_worksheets.Values).GetEnumerator();
        }

        #endregion

        #region IXLWorksheets Members

        public int Count
        {
            [DebuggerStepThrough]
            get { return _worksheets.Count; }
        }

        public bool TryGetWorksheet(string sheetName, out IXLWorksheet worksheet)
        {
            XLWorksheet w;
            if (_worksheets.TryGetValue(sheetName, out w))
            {
                worksheet = w;
                return true;
            }
            worksheet = null;
            return false;
        }

        public IXLWorksheet Worksheet(String sheetName)
        {
            XLWorksheet w;
            if (_worksheets.TryGetValue(sheetName, out w))
                return w;

            var wss = _worksheets.Where(ws => ws.Key.ToLower().Equals(sheetName.ToLower()));

            if (wss.Any())
                return wss.First().Value;

            throw new Exception("There isn't a worksheet named '" + sheetName + "'.");
        }

        public IXLWorksheet Worksheet(Int32 position)
        {
            int wsCount = _worksheets.Values.Count(w => w.Position == position);
            if (wsCount == 0)
                throw new Exception("There isn't a worksheet associated with that position.");

            if (wsCount > 1)
            {
                throw new Exception(
                    "Can't retrieve a worksheet because there are multiple worksheets associated with that position.");
            }

            return _worksheets.Values.Single(w => w.Position == position);
        }

        public IXLWorksheet Add(String sheetName)
        {
            var sheet = new XLWorksheet(sheetName, _workbook);
            _worksheets.Add(sheetName, sheet);
            sheet._position = _worksheets.Count + _workbook.UnsupportedSheets.Count;
            return sheet;
        }

        public IXLWorksheet Add(String sheetName, Int32 position)
        {
            _worksheets.Values.Where(w => w._position >= position).ForEach(w => w._position += 1);
            _workbook.UnsupportedSheets.Where(w => w.Position >= position).ForEach(w => w.Position += 1);
            var sheet = new XLWorksheet(sheetName, _workbook);
            _worksheets.Add(sheetName, sheet);
            sheet._position = position;
            return sheet;
        }

        public void Delete(String sheetName)
        {
            Delete(_worksheets[sheetName].Position);
        }

        public void Delete(Int32 position)
        {
            int wsCount = _worksheets.Values.Count(w => w.Position == position);
            if (wsCount == 0)
                throw new Exception("There isn't a worksheet associated with that index.");

            if (wsCount > 1)
                throw new Exception(
                    "Can't delete the worksheet because there are multiple worksheets associated with that index.");

            var ws = _worksheets.Values.Single(w => w.Position == position);
            if (!XLHelper.IsNullOrWhiteSpace(ws.RelId) && !Deleted.Contains(ws.RelId))
                Deleted.Add(ws.RelId);

            _worksheets.RemoveAll(w => w.Position == position);
            _worksheets.Values.Where(w => w.Position > position).ForEach(w => w._position -= 1);
            _workbook.UnsupportedSheets.Where(w => w.Position > position).ForEach(w => w.Position -= 1);
        }

        IEnumerator<IXLWorksheet> IEnumerable<IXLWorksheet>.GetEnumerator()
        {
            return _worksheets.Values.Cast<IXLWorksheet>().GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public IXLWorksheet Add(DataTable dataTable)
        {
            return Add(dataTable, dataTable.TableName);
        }

        public IXLWorksheet Add(DataTable dataTable, String sheetName)
        {
            var ws = Add(sheetName);
            ws.Cell(1, 1).InsertTable(dataTable);
            ws.Columns().AdjustToContents(1, 75);
            return ws;
        }

        public void Add(DataSet dataSet)
        {
            foreach (DataTable t in dataSet.Tables)
                Add(t);
        }

        #endregion

        public void Rename(String oldSheetName, String newSheetName)
        {
            if (XLHelper.IsNullOrWhiteSpace(oldSheetName) || !_worksheets.ContainsKey(oldSheetName)) return;

            var ws = _worksheets[oldSheetName];
            _worksheets.Remove(oldSheetName);
            _worksheets.Add(newSheetName, ws);
        }
    }
}