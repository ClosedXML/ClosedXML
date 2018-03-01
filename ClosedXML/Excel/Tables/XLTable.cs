using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLTable : XLRange, IXLTable
    {
        #region Private fields

        private string _name;
        internal bool _showTotalsRow;
        internal HashSet<String> _uniqueNames;

        #endregion Private fields

        #region Constructor

        public XLTable(XLRange range, Boolean addToTables, Boolean setAutofilter = true)
            : base(new XLRangeParameters(range.RangeAddress, range.Style))
        {
            InitializeValues(setAutofilter);

            Int32 id = 1;
            while (true)
            {
                string tableName = String.Format("Table{0}", id);
                if (!Worksheet.Tables.Any(t => t.Name == tableName))
                {
                    Name = tableName;
                    AddToTables(range, addToTables);
                    break;
                }
                id++;
            }
        }

        public XLTable(XLRange range, String name, Boolean addToTables, Boolean setAutofilter = true)
            : base(new XLRangeParameters(range.RangeAddress, range.Style))
        {
            InitializeValues(setAutofilter);

            Name = name;
            AddToTables(range, addToTables);
        }

        #endregion Constructor

        private IXLRangeAddress _lastRangeAddress;
        private Dictionary<String, IXLTableField> _fieldNames = null;

        public Dictionary<String, IXLTableField> FieldNames
        {
            get
            {
                if (_fieldNames != null && _lastRangeAddress != null && _lastRangeAddress.Equals(RangeAddress))
                    return _fieldNames;

                if (_fieldNames == null)
                {
                    _fieldNames = new Dictionary<String, IXLTableField>();
                    _lastRangeAddress = RangeAddress;
                    HeadersRow();
                }
                else
                {
                    HeadersRow(false);
                }

                RescanFieldNames();

                _lastRangeAddress = RangeAddress;

                return _fieldNames;
            }
        }

        private void RescanFieldNames()
        {
            if (ShowHeaderRow)
            {
                var detectedFieldNames = new Dictionary<String, IXLTableField>();
                var headersRow = HeadersRow(false);
                Int32 cellPos = 0;
                foreach (var cell in headersRow.Cells())
                {
                    var name = cell.GetString();
                    if (_fieldNames.ContainsKey(name) && _fieldNames[name].Column.ColumnNumber() == cell.Address.ColumnNumber)
                    {
                        (_fieldNames[name] as XLTableField).Index = cellPos;
                        detectedFieldNames.Add(name, _fieldNames[name]);
                        cellPos++;
                        continue;
                    }

                    if (String.IsNullOrWhiteSpace(name))
                    {
                        name = GetUniqueName("Column", cellPos + 1, true);
                        cell.SetValue(name);
                        cell.DataType = XLDataType.Text;
                    }
                    if (_fieldNames.ContainsKey(name))
                        throw new ArgumentException("The header row contains more than one field name '" + name + "'.");

                    _fieldNames.Add(name, new XLTableField(this, name) { Index = cellPos++ });
                    detectedFieldNames.Add(name, _fieldNames[name]);
                }

                _fieldNames.Keys
                    .Where(key => !detectedFieldNames.ContainsKey(key))
                    .ToArray()
                    .ForEach(key => _fieldNames.Remove(key));
            }
            else
            {
                Int32 colCount = ColumnCount();
                for (Int32 i = 1; i <= colCount; i++)
                {
                    if (!_fieldNames.Values.Any(f => f.Index == i - 1))
                    {
                        var name = "Column" + i;

                        _fieldNames.Add(name, new XLTableField(this, name) { Index = i - 1 });
                    }
                }
            }
        }

        internal void AddFields(IEnumerable<String> fieldNames)
        {
            _fieldNames = new Dictionary<String, IXLTableField>();

            Int32 cellPos = 0;
            foreach (var name in fieldNames)
            {
                _fieldNames.Add(name, new XLTableField(this, name) { Index = cellPos++ });
            }
        }

        internal void RenameField(String oldName, String newName)
        {
            if (!_fieldNames.ContainsKey(oldName))
                throw new ArgumentException("The field does not exist in this table", "oldName");

            var field = _fieldNames[oldName];
            _fieldNames.Remove(oldName);
            _fieldNames.Add(newName, field);
        }

        internal String RelId { get; set; }

        public IXLTableRange DataRange
        {
            get
            {
                XLRange range;

                if (_showHeaderRow)
                {
                    range = _showTotalsRow
                                ? Range(2, 1, RowCount() - 1, ColumnCount())
                                : Range(2, 1, RowCount(), ColumnCount());
                }
                else
                {
                    range = _showTotalsRow
                                ? Range(1, 1, RowCount() - 1, ColumnCount())
                                : Range(1, 1, RowCount(), ColumnCount());
                }

                return new XLTableRange(range, this);
            }
        }

        private XLAutoFilter _autoFilter;

        public XLAutoFilter AutoFilter
        {
            get
            {
                using (var asRange = ShowTotalsRow ? Range(1, 1, RowCount() - 1, ColumnCount()) : AsRange())
                {
                    if (_autoFilter == null)
                        _autoFilter = new XLAutoFilter();

                    _autoFilter.Range = asRange;
                }
                return _autoFilter;
            }
        }

        public new IXLBaseAutoFilter SetAutoFilter()
        {
            return AutoFilter;
        }

        #region IXLTable Members

        public Boolean EmphasizeFirstColumn { get; set; }
        public Boolean EmphasizeLastColumn { get; set; }
        public Boolean ShowRowStripes { get; set; }
        public Boolean ShowColumnStripes { get; set; }

        private Boolean _showAutoFilter;

        public Boolean ShowAutoFilter
        {
            get { return _showHeaderRow && _showAutoFilter; }
            set { _showAutoFilter = value; }
        }

        public XLTableTheme Theme { get; set; }

        public String Name
        {
            get { return _name; }
            set
            {
                if (Worksheet.Tables.Any(t => t.Name == value))
                {
                    throw new ArgumentException(String.Format("This worksheet already contains a table named '{0}'",
                                                              value));
                }

                _name = value;
            }
        }

        public Boolean ShowTotalsRow
        {
            get { return _showTotalsRow; }
            set
            {
                if (value && !_showTotalsRow)
                    InsertRowsBelow(1);
                else if (!value && _showTotalsRow)
                    TotalsRow().Delete();

                _showTotalsRow = value;

                if (_showTotalsRow)
                {
                    AutoFilter.Range = Worksheet.Range(
                        RangeAddress.FirstAddress.RowNumber, RangeAddress.FirstAddress.ColumnNumber,
                        RangeAddress.LastAddress.RowNumber - 1, RangeAddress.LastAddress.ColumnNumber);
                }
                else
                    AutoFilter.Range = Worksheet.Range(RangeAddress);
            }
        }

        public IXLRangeRow HeadersRow()
        {
            return HeadersRow(true);
        }

        internal IXLRangeRow HeadersRow(Boolean scanForNewFieldsNames)
        {
            if (!ShowHeaderRow) return null;

            if (scanForNewFieldsNames)
            {
                var tempResult = FieldNames;
            }

            return FirstRow();
        }

        public IXLRangeRow TotalsRow()
        {
            return ShowTotalsRow ? LastRow() : null;
        }

        public IXLTableField Field(String fieldName)
        {
            return Field(GetFieldIndex(fieldName));
        }

        public IXLTableField Field(Int32 fieldIndex)
        {
            return FieldNames.Values.First(f => f.Index == fieldIndex);
        }

        public IEnumerable<IXLTableField> Fields
        {
            get
            {
                Int32 columnCount = ColumnCount();
                for (int co = 0; co < columnCount; co++)
                    yield return Field(co);
            }
        }

        public IXLTable Resize(IXLRangeAddress rangeAddress)
        {
            return Resize(Worksheet.Range(RangeAddress));
        }

        public IXLTable Resize(string rangeAddress)
        {
            return Resize(Worksheet.Range(RangeAddress));
        }

        public IXLTable Resize(IXLCell firstCell, IXLCell lastCell)
        {
            return Resize(Worksheet.Range(firstCell, lastCell));
        }

        public IXLTable Resize(string firstCellAddress, string lastCellAddress)
        {
            return Resize(Worksheet.Range(firstCellAddress, lastCellAddress));
        }

        public IXLTable Resize(IXLAddress firstCellAddress, IXLAddress lastCellAddress)
        {
            return Resize(Worksheet.Range(firstCellAddress, lastCellAddress));
        }

        public IXLTable Resize(int firstCellRow, int firstCellColumn, int lastCellRow, int lastCellColumn)
        {
            return Resize(Worksheet.Range(firstCellRow, firstCellColumn, lastCellRow, lastCellColumn));
        }

        public IXLTable Resize(IXLRange range)
        {
            if (!this.ShowHeaderRow)
                throw new NotImplementedException("Resizing of tables with no headers not supported yet.");

            if (this.Worksheet != range.Worksheet)
                throw new InvalidOperationException("You cannot resize a table to a range on a different sheet.");

            var totalsRowChanged = this.ShowTotalsRow ? range.LastRow().RowNumber() - this.TotalsRow().RowNumber() : 0;
            var oldTotalsRowNumber = this.ShowTotalsRow ? this.TotalsRow().RowNumber() : -1;

            var existingHeaders = this.FieldNames.Keys;
            var newHeaders = new HashSet<string>();
            var tempArray = this.Fields.Select(f => f.Column).ToArray();

            var firstRow = range.Row(1);
            if (!firstRow.FirstCell().Address.Equals(this.HeadersRow().FirstCell().Address)
                || !firstRow.LastCell().Address.Equals(this.HeadersRow().LastCell().Address))
            {
                _uniqueNames.Clear();
                var co = 1;
                foreach (var c in firstRow.Cells())
                {
                    if (String.IsNullOrWhiteSpace(((XLCell)c).InnerText))
                        c.Value = GetUniqueName("Column", co, true);

                    var header = c.GetString();
                    _uniqueNames.Add(header);

                    if (!existingHeaders.Contains(header))
                        newHeaders.Add(header);

                    co++;
                }
            }

            if (totalsRowChanged < 0)
            {
                range.Rows(r => r.RowNumber().Equals(this.TotalsRow().RowNumber() + totalsRowChanged)).Single().InsertRowsAbove(1);
                range = Worksheet.Range(range.FirstCell(), range.LastCell().CellAbove());
                oldTotalsRowNumber++;
            }
            else if (totalsRowChanged > 0)
            {
                this.TotalsRow().RowBelow(totalsRowChanged + 1).InsertRowsAbove(1);
                this.TotalsRow().AsRange().Delete(XLShiftDeletedCells.ShiftCellsUp);
            }

            this.RangeAddress = range.RangeAddress as XLRangeAddress;
            RescanFieldNames();

            if (this.ShowTotalsRow)
            {
                foreach (var f in this._fieldNames.Values)
                {
                    var c = this.TotalsRow().Cell(f.Index + 1);
                    if (!c.IsEmpty() && newHeaders.Contains(f.Name))
                    {
                        f.TotalsRowLabel = c.GetFormattedString();
                        c.DataType = XLDataType.Text;
                    }
                }

                if (totalsRowChanged != 0)
                {
                    foreach (var f in this._fieldNames.Values.Cast<XLTableField>())
                    {
                        f.UpdateUnderlyingCellFormula();
                        var c = this.TotalsRow().Cell(f.Index + 1);
                        if (!String.IsNullOrWhiteSpace(f.TotalsRowLabel))
                        {
                            c.DataType = XLDataType.Text;

                            //Remove previous row's label
                            var oldTotalsCell = this.Worksheet.Cell(oldTotalsRowNumber, f.Column.ColumnNumber());
                            if (oldTotalsCell.Value.ToString() == f.TotalsRowLabel)
                                oldTotalsCell.Value = null;
                        }

                        if (f.TotalsRowFunction != XLTotalsRowFunction.None)
                            c.DataType = XLDataType.Number;
                    }
                }
            }

            return this;
        }

        public IXLTable SetEmphasizeFirstColumn()
        {
            EmphasizeFirstColumn = true;
            return this;
        }

        public IXLTable SetEmphasizeFirstColumn(Boolean value)
        {
            EmphasizeFirstColumn = value;
            return this;
        }

        public IXLTable SetEmphasizeLastColumn()
        {
            EmphasizeLastColumn = true;
            return this;
        }

        public IXLTable SetEmphasizeLastColumn(Boolean value)
        {
            EmphasizeLastColumn = value;
            return this;
        }

        public IXLTable SetShowRowStripes()
        {
            ShowRowStripes = true;
            return this;
        }

        public IXLTable SetShowRowStripes(Boolean value)
        {
            ShowRowStripes = value;
            return this;
        }

        public IXLTable SetShowColumnStripes()
        {
            ShowColumnStripes = true;
            return this;
        }

        public IXLTable SetShowColumnStripes(Boolean value)
        {
            ShowColumnStripes = value;
            return this;
        }

        public IXLTable SetShowTotalsRow()
        {
            ShowTotalsRow = true;
            return this;
        }

        public IXLTable SetShowTotalsRow(Boolean value)
        {
            ShowTotalsRow = value;
            return this;
        }

        public IXLTable SetShowAutoFilter()
        {
            ShowAutoFilter = true;
            return this;
        }

        public IXLTable SetShowAutoFilter(Boolean value)
        {
            ShowAutoFilter = value;
            return this;
        }

        public new IXLRange Sort(String columnsToSortBy, XLSortOrder sortOrder = XLSortOrder.Ascending,
                                 Boolean matchCase = false, Boolean ignoreBlanks = true)
        {
            var toSortBy = new StringBuilder();
            foreach (string coPairTrimmed in columnsToSortBy.Split(',').Select(coPair => coPair.Trim()))
            {
                String coString;
                String order;
                if (coPairTrimmed.Contains(' '))
                {
                    var pair = coPairTrimmed.Split(' ');
                    coString = pair[0];
                    order = pair[1];
                }
                else
                {
                    coString = coPairTrimmed;
                    order = sortOrder == XLSortOrder.Ascending ? "ASC" : "DESC";
                }

                Int32 co;
                if (!Int32.TryParse(coString, out co))
                    co = Field(coString).Index + 1;

                if (toSortBy.Length > 0)
                    toSortBy.Append(',');

                toSortBy.Append(co);
                toSortBy.Append(' ');
                toSortBy.Append(order);
            }
            return DataRange.Sort(toSortBy.ToString(), sortOrder, matchCase, ignoreBlanks);
        }

        public new IXLTable Clear(XLClearOptions clearOptions = XLClearOptions.ContentsAndFormats)
        {
            base.Clear(clearOptions);
            return this;
        }

        IXLBaseAutoFilter IXLTable.AutoFilter
        {
            get { return AutoFilter; }
        }

        public new void Dispose()
        {
            if (AutoFilter != null)
                AutoFilter.Dispose();

            base.Dispose();
        }

        #endregion IXLTable Members

        private void InitializeValues(Boolean setAutofilter)
        {
            ShowRowStripes = true;
            _showHeaderRow = true;
            Theme = XLTableTheme.TableStyleLight9;
            if (setAutofilter)
                InitializeAutoFilter();

            AsRange().Row(1).DataType = XLDataType.Text;

            if (RowCount() == 1)
                InsertRowsBelow(1);
        }

        public void InitializeAutoFilter()
        {
            ShowAutoFilter = true;
        }

        private void AddToTables(XLRange range, Boolean addToTables)
        {
            if (!addToTables) return;

            _uniqueNames = new HashSet<string>();
            Int32 co = 1;
            foreach (IXLCell c in range.Row(1).Cells())
            {
                if (String.IsNullOrWhiteSpace(((XLCell)c).InnerText))
                    c.Value = GetUniqueName("Column", co, true);
                _uniqueNames.Add(c.GetString());
                co++;
            }
            Worksheet.Tables.Add(this);
        }

        private String GetUniqueName(String originalName, Int32 initialOffset, Boolean enforceOffset)
        {
            String name = String.Concat(originalName, enforceOffset ? initialOffset.ToInvariantString() : string.Empty);
            if (_uniqueNames?.Contains(name) ?? false)
            {
                Int32 i = initialOffset;
                name = originalName + i.ToInvariantString();
                while (_uniqueNames.Contains(name))
                {
                    i++;
                    name = originalName + i.ToInvariantString();
                }
            }

            return name;
        }

        public Int32 GetFieldIndex(String name)
        {
            // There is a discrepancy in the way headers with line breaks are stored.
            // The entry in the table definition will contain \r\n
            // but the shared string value of the actual cell will contain only \n
            name = name.Replace("\r\n", "\n");
            if (FieldNames.ContainsKey(name))
                return FieldNames[name].Index;

            throw new ArgumentOutOfRangeException("The header row doesn't contain field name '" + name + "'.");
        }

        internal Boolean _showHeaderRow;

        public Boolean ShowHeaderRow
        {
            get { return _showHeaderRow; }
            set
            {
                if (_showHeaderRow == value) return;

                if (_showHeaderRow)
                {
                    var headersRow = HeadersRow();
                    _uniqueNames = new HashSet<string>();
                    Int32 co = 1;
                    foreach (IXLCell c in headersRow.Cells())
                    {
                        if (String.IsNullOrWhiteSpace(((XLCell)c).InnerText))
                            c.Value = GetUniqueName("Column", co, true);
                        _uniqueNames.Add(c.GetString());
                        co++;
                    }

                    headersRow.Clear();
                    RangeAddress.FirstAddress = new XLAddress(Worksheet, RangeAddress.FirstAddress.RowNumber + 1,
                                          RangeAddress.FirstAddress.ColumnNumber,
                                          RangeAddress.FirstAddress.FixedRow,
                                          RangeAddress.FirstAddress.FixedColumn);
                }
                else
                {
                    using (var asRange = Worksheet.Range(
                        RangeAddress.FirstAddress.RowNumber - 1,
                        RangeAddress.FirstAddress.ColumnNumber,
                        RangeAddress.LastAddress.RowNumber,
                        RangeAddress.LastAddress.ColumnNumber
                        ))
                    using (var firstRow = asRange.FirstRow())
                    {
                        IXLRangeRow rangeRow;
                        if (firstRow.IsEmpty(true))
                        {
                            rangeRow = firstRow;
                            RangeAddress.FirstAddress = new XLAddress(Worksheet,
                                  RangeAddress.FirstAddress.RowNumber - 1,
                                  RangeAddress.FirstAddress.ColumnNumber,
                                  RangeAddress.FirstAddress.FixedRow,
                                  RangeAddress.FirstAddress.FixedColumn);
                        }
                        else
                        {
                            var fAddress = RangeAddress.FirstAddress;
                            var lAddress = RangeAddress.LastAddress;

                            rangeRow = firstRow.InsertRowsBelow(1, false).First();

                            RangeAddress.FirstAddress = new XLAddress(Worksheet, fAddress.RowNumber,
                                                                      fAddress.ColumnNumber,
                                                                      fAddress.FixedRow,
                                                                      fAddress.FixedColumn);

                            RangeAddress.LastAddress = new XLAddress(Worksheet, lAddress.RowNumber + 1,
                                                                     lAddress.ColumnNumber,
                                                                     lAddress.FixedRow,
                                                                     lAddress.FixedColumn);
                        }

                        Int32 co = 1;
                        foreach (var name in FieldNames.Values.Select(f => f.Name))
                        {
                            rangeRow.Cell(co).SetValue(name);
                            co++;
                        }
                    }
                }
                _showHeaderRow = value;

                if (_showHeaderRow)
                    HeadersRow().DataType = XLDataType.Text;
            }
        }

        public IXLTable SetShowHeaderRow()
        {
            return SetShowHeaderRow(true);
        }

        public IXLTable SetShowHeaderRow(Boolean value)
        {
            ShowHeaderRow = value;
            return this;
        }

        public void ExpandTableRows(Int32 rows)
        {
            RangeAddress.LastAddress = new XLAddress(Worksheet, RangeAddress.LastAddress.RowNumber + rows,
                                                     RangeAddress.LastAddress.ColumnNumber,
                                                     RangeAddress.LastAddress.FixedRow,
                                                     RangeAddress.LastAddress.FixedColumn);
        }

        public override XLRangeColumn Column(int columnNumber)
        {
            var column = base.Column(columnNumber);
            column.Table = this;
            return column;
        }

        public override XLRangeColumn Column(string columnName)
        {
            var column = base.Column(columnName);
            column.Table = this;
            return column;
        }

        public override IXLRangeColumns Columns(int firstColumn, int lastColumn)
        {
            var columns = base.Columns(firstColumn, lastColumn);
            columns.Cast<XLRangeColumn>().ForEach(column => column.Table = this);
            return columns;
        }

        public override IXLRangeColumns Columns(Func<IXLRangeColumn, bool> predicate = null)
        {
            var columns = base.Columns(predicate);
            columns.Cast<XLRangeColumn>().ForEach(column => column.Table = this);
            return columns;
        }

        public override IXLRangeColumns Columns(string columns)
        {
            var cols = base.Columns(columns);
            cols.Cast<XLRangeColumn>().ForEach(column => column.Table = this);
            return cols;
        }

        public override IXLRangeColumns Columns(string firstColumn, string lastColumn)
        {
            var columns = base.Columns(firstColumn, lastColumn);
            columns.Cast<XLRangeColumn>().ForEach(column => column.Table = this);
            return columns;
        }

        public override XLRangeColumns ColumnsUsed(bool includeFormats, Func<IXLRangeColumn, bool> predicate = null)
        {
            var columns = base.ColumnsUsed(includeFormats, predicate);
            columns.Cast<XLRangeColumn>().ForEach(column => column.Table = this);
            return columns;
        }

        public override XLRangeColumns ColumnsUsed(Func<IXLRangeColumn, bool> predicate = null)
        {
            var columns = base.ColumnsUsed(predicate);
            columns.Cast<XLRangeColumn>().ForEach(column => column.Table = this);
            return columns;
        }

        public IEnumerable<dynamic> AsDynamicEnumerable()
        {
            foreach (var row in this.DataRange.Rows())
            {
                dynamic expando = new ExpandoObject();
                foreach (var f in this.Fields)
                {
                    var value = row.Cell(f.Index + 1).Value;
                    // ExpandoObject supports IDictionary so we can extend it like this
                    var expandoDict = expando as IDictionary<string, object>;
                    if (expandoDict.ContainsKey(f.Name))
                        expandoDict[f.Name] = value;
                    else
                        expandoDict.Add(f.Name, value);
                }

                yield return expando;
            }
        }

        public DataTable AsNativeDataTable()
        {
            var table = new DataTable(this.Name);

            foreach (var f in Fields.Cast<XLTableField>())
            {
                Type type = typeof(object);
                if (f.IsConsistentDataType())
                {
                    var c = f.Column.Cells().Skip(this.ShowHeaderRow ? 1 : 0).First();
                    switch (c.DataType)
                    {
                        case XLDataType.Text:
                            type = typeof(String);
                            break;

                        case XLDataType.Boolean:
                            type = typeof(Boolean);
                            break;

                        case XLDataType.DateTime:
                            type = typeof(DateTime);
                            break;

                        case XLDataType.TimeSpan:
                            type = typeof(TimeSpan);
                            break;

                        case XLDataType.Number:
                            type = typeof(Double);
                            break;
                    }
                }

                table.Columns.Add(f.Name, type);
            }

            foreach (var row in this.DataRange.Rows())
            {
                var dr = table.NewRow();

                foreach (var f in this.Fields)
                {
                    dr[f.Name] = row.Cell(f.Index + 1).Value;
                }

                table.Rows.Add(dr);
            }

            return table;
        }
    }
}
