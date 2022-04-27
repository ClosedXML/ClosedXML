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
        private readonly XLWorkbook _workbook;
        private readonly Dictionary<string, XLWorksheet> _worksheets = new Dictionary<string, XLWorksheet>(StringComparer.OrdinalIgnoreCase);
        internal ICollection<string> Deleted { get; private set; }

        #region Constructor

        public XLWorksheets(XLWorkbook workbook)
        {
            _workbook = workbook;
            Deleted = new HashSet<string>();
        }

        #endregion Constructor

        #region IEnumerable<XLWorksheet> Members

        public IEnumerator<XLWorksheet> GetEnumerator()
        {
            return ((IEnumerable<XLWorksheet>)_worksheets.Values).GetEnumerator();
        }

        #endregion IEnumerable<XLWorksheet> Members

        #region IXLWorksheets Members

        public int Count
        {
            [DebuggerStepThrough]
            get { return _worksheets.Count; }
        }

        public bool Contains(string sheetName)
        {
            return _worksheets.ContainsKey(sheetName);
        }

        public bool TryGetWorksheet(string sheetName, out IXLWorksheet worksheet)
        {
            if (_worksheets.TryGetValue(sheetName.UnescapeSheetName(), out var w))
            {
                worksheet = w;
                return true;
            }
            worksheet = null;
            return false;
        }

        public IXLWorksheet Worksheet(string sheetName)
        {
            sheetName = sheetName.UnescapeSheetName();

            if (_worksheets.TryGetValue(sheetName, out var w))
            {
                return w;
            }

            throw new ArgumentException("There isn't a worksheet named '" + sheetName + "'.");
        }

        public IXLWorksheet Worksheet(int position)
        {
            var wsCount = _worksheets.Values.Count(w => w.Position == position);
            if (wsCount == 0)
            {
                throw new ArgumentException("There isn't a worksheet associated with that position.");
            }

            if (wsCount > 1)
            {
                throw new ArgumentException(
                    "Can't retrieve a worksheet because there are multiple worksheets associated with that position.");
            }

            return _worksheets.Values.Single(w => w.Position == position);
        }

        public IXLWorksheet Add()
        {
            return Add(GetNextWorksheetName());
        }

        public IXLWorksheet Add(int position)
        {
            return Add(GetNextWorksheetName(), position);
        }

        public IXLWorksheet Add(string sheetName)
        {
            var sheet = new XLWorksheet(sheetName, _workbook);
            Add(sheetName, sheet);
            sheet._position = _worksheets.Count + _workbook.UnsupportedSheets.Count;
            return sheet;
        }

        public IXLWorksheet Add(string sheetName, int position)
        {
            _worksheets.Values.Where(w => w._position >= position).ForEach(w => w._position += 1);
            _workbook.UnsupportedSheets.Where(w => w.Position >= position).ForEach(w => w.Position += 1);
            var sheet = new XLWorksheet(sheetName, _workbook);
            Add(sheetName, sheet);
            sheet._position = position;
            return sheet;
        }

        private void Add(string sheetName, XLWorksheet sheet)
        {
            if (_worksheets.ContainsKey(sheetName))
            {
                throw new ArgumentException(string.Format("A worksheet with the same name ({0}) has already been added.", sheetName), nameof(sheetName));
            }

            _worksheets.Add(sheetName, sheet);
        }

        public void Delete(string sheetName)
        {
            Delete(_worksheets[sheetName].Position);
        }

        public void Delete(int position)
        {
            var wsCount = _worksheets.Values.Count(w => w.Position == position);
            if (wsCount == 0)
            {
                throw new ArgumentException("There isn't a worksheet associated with that index.");
            }

            if (wsCount > 1)
            {
                throw new ArgumentException(
                    "Can't delete the worksheet because there are multiple worksheets associated with that index.");
            }

            var ws = _worksheets.Values.Single(w => w.Position == position);
            if (!string.IsNullOrWhiteSpace(ws.RelId) && !Deleted.Contains(ws.RelId))
            {
                Deleted.Add(ws.RelId);
            }

            _worksheets.RemoveAll(w => w.Position == position);
            _worksheets.Values.Where(w => w.Position > position).ForEach(w => w._position -= 1);
            _workbook.UnsupportedSheets.Where(w => w.Position > position).ForEach(w => w.Position -= 1);
            _workbook.InvalidateFormulas();

            ws.Cleanup();
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

        public IXLWorksheet Add(DataTable dataTable, string sheetName)
        {
            var ws = Add(sheetName);
            ws.Cell(1, 1).InsertTable(dataTable, sheetName);
            return ws;
        }

        public void Add(DataSet dataSet)
        {
            foreach (DataTable t in dataSet.Tables)
            {
                Add(t);
            }
        }

        #endregion IXLWorksheets Members

        public void Rename(string oldSheetName, string newSheetName)
        {
            if (string.IsNullOrWhiteSpace(oldSheetName) || !_worksheets.TryGetValue(oldSheetName, out var ws))
            {
                return;
            }

            if (!oldSheetName.Equals(newSheetName, StringComparison.OrdinalIgnoreCase)
                && _worksheets.ContainsKey(newSheetName))
            {
                throw new ArgumentException(string.Format("A worksheet with the same name ({0}) has already been added.", newSheetName), nameof(newSheetName));
            }

            _worksheets.Remove(oldSheetName);
            Add(newSheetName, ws);
        }

        #region Private members

        private string GetNextWorksheetName()
        {
            var worksheetNumber = Count + 1;
            var sheetName = $"Sheet{worksheetNumber}";
            while (_worksheets.Values.Any(p => p.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase)))
            {
                worksheetNumber++;
                sheetName = $"Sheet{worksheetNumber}";
            }
            return sheetName;
        }

        #endregion Private members
    }
}
