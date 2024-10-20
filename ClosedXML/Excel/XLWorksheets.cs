using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLWorksheets : IXLWorksheets, IEnumerable<XLWorksheet>
    {
        private readonly XLWorkbook _workbook;
        private readonly Dictionary<String, XLWorksheet> _worksheets = new Dictionary<String, XLWorksheet>(StringComparer.OrdinalIgnoreCase);
        internal ICollection<String> Deleted { get; private set; }

        /// <summary>
        /// SheetId that will be assigned to next created sheet.
        /// </summary>
        private UInt32 _nextSheetId = 1;

        #region Constructor

        public XLWorksheets(XLWorkbook workbook)
        {
            _workbook = workbook;
            Deleted = new HashSet<String>();
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

        public Boolean Contains(String sheetName)
        {
            return _worksheets.ContainsKey(sheetName);
        }

        bool IXLWorksheets.TryGetWorksheet(string sheetName, [NotNullWhen(true)] out IXLWorksheet? worksheet)
        {
            if (TryGetWorksheet(sheetName, out var foundSheet))
            {
                worksheet = foundSheet;
                return true;
            }

            worksheet = null;
            return false;
        }

        internal bool TryGetWorksheet(string sheetName, [NotNullWhen(true)] out XLWorksheet? worksheet)
        {
            if (_worksheets.TryGetValue(sheetName.UnescapeSheetName(), out worksheet))
            {
                return true;
            }

            worksheet = null;
            return false;
        }

        public IXLWorksheet Worksheet(String sheetName)
        {
            sheetName = sheetName.UnescapeSheetName();

            if (_worksheets.TryGetValue(sheetName, out XLWorksheet w))
                return w;

            throw new ArgumentException("There isn't a worksheet named '" + sheetName + "'.");
        }

        public IXLWorksheet Worksheet(Int32 position)
        {
            int wsCount = _worksheets.Values.Count(w => w.Position == position);
            if (wsCount == 0)
                throw new ArgumentException("There isn't a worksheet associated with that position.");

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

        public IXLWorksheet Add(Int32 position)
        {
            return Add(GetNextWorksheetName(), position);
        }

        public IXLWorksheet Add(String sheetName)
        {
            var sheet = new XLWorksheet(sheetName, _workbook, GetNextSheetId());
            Add(sheetName, sheet);
            sheet._position = _worksheets.Count + _workbook.UnsupportedSheets.Count;
            return sheet;
        }

        public IXLWorksheet Add(String sheetName, Int32 position)
        {
            return Add(sheetName, position, GetNextSheetId());
        }

        internal XLWorksheet Add(String sheetName, Int32 position, UInt32 sheetId)
        {
            _worksheets.Values.Where(w => w._position >= position).ForEach(w => w._position += 1);
            _workbook.UnsupportedSheets.Where(w => w.Position >= position).ForEach(w => w.Position += 1);

            // If the loaded sheetId is greater than current, just make sure our next sheetId is even bigger.
            _nextSheetId = Math.Max(_nextSheetId, sheetId + 1);
            var sheet = new XLWorksheet(sheetName, _workbook, sheetId);
            Add(sheetName, sheet);
            sheet._position = position;
            return sheet;
        }

        private void Add(String sheetName, XLWorksheet sheet)
        {
            if (_worksheets.ContainsKey(sheetName))
                throw new ArgumentException(String.Format("A worksheet with the same name ({0}) has already been added.", sheetName), nameof(sheetName));

            _worksheets.Add(sheetName, sheet);

            _workbook.NotifyWorksheetAdded(sheet);
        }

        public void Delete(String sheetName)
        {
            Delete(_worksheets[sheetName].Position);
        }

        public void Delete(Int32 position)
        {
            int wsCount = _worksheets.Values.Count(w => w.Position == position);
            if (wsCount == 0)
                throw new ArgumentException("There isn't a worksheet associated with that index.");

            if (wsCount > 1)
                throw new ArgumentException(
                    "Can't delete the worksheet because there are multiple worksheets associated with that index.");

            var ws = _worksheets.Values.Single(w => w.Position == position);
            if (!String.IsNullOrWhiteSpace(ws.RelId) && !Deleted.Contains(ws.RelId))
                Deleted.Add(ws.RelId);

            _worksheets.RemoveAll(w => w.Position == position);
            _worksheets.Values.Where(w => w.Position > position).ForEach(w => w._position -= 1);
            _workbook.UnsupportedSheets.Where(w => w.Position > position).ForEach(w => w.Position -= 1);

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

        public IXLWorksheet Add(DataTable dataTable, String sheetName)
        {
            return Add(dataTable, sheetName, TableNameGenerator.GetNewTableName(_workbook));
        }

        public IXLWorksheet Add(DataTable dataTable, String sheetName, string tableName)
        {
            var ws = Add(sheetName);
            ws.Cell(1, 1).InsertTable(dataTable, tableName);
            return ws;
        }

        public void Add(DataSet dataSet)
        {
            foreach (DataTable t in dataSet.Tables)
                Add(t);
        }

        #endregion IXLWorksheets Members

        public void Rename(String oldSheetName, String newSheetName)
        {
            if (String.IsNullOrWhiteSpace(oldSheetName) || !_worksheets.TryGetValue(oldSheetName, out XLWorksheet ws)) return;

            if (!oldSheetName.Equals(newSheetName, StringComparison.OrdinalIgnoreCase)
                && _worksheets.ContainsKey(newSheetName))
                throw new ArgumentException(String.Format("A worksheet with the same name ({0}) has already been added.", newSheetName), nameof(newSheetName));

            _worksheets.Remove(oldSheetName);
            Add(newSheetName, ws);

            foreach (var listener in GetWorkbookListeners())
                listener.OnSheetRenamed(oldSheetName, newSheetName);
        }

        #region Private members

        private IEnumerable<IWorkbookListener> GetWorkbookListeners()
        {
            // All components that should be updated when sheet is added/removed or renamed should
            // be enumerated here.
            yield return _workbook.CalcEngine;

            foreach (var sheet in _worksheets.Values)
            {
                yield return sheet.Internals.CellsCollection;
            }

            foreach (var definedName in _workbook.DefinedNamesInternal)
                yield return definedName;

            foreach (var sheet in _worksheets.Values)
            {
                foreach (var definedName in sheet.DefinedNames)
                {
                    yield return definedName;
                }

                yield return sheet.PageSetupInternal.PrintAreasInternal;
            }
        }

        private String GetNextWorksheetName()
        {
            var worksheetNumber = this.Count + 1;
            var sheetName = $"Sheet{worksheetNumber}";
            while (_worksheets.Values.Any(p => p.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase)))
            {
                worksheetNumber++;
                sheetName = $"Sheet{worksheetNumber}";
            }
            return sheetName;
        }

        private UInt32 GetNextSheetId() => _nextSheetId++;

        #endregion Private members
    }
}
