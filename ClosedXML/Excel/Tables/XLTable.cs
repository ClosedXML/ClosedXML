using System;
using System.Collections.Generic;
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

        #endregion

        #region Constructor

        public XLTable(XLRange range, Boolean addToTables, Boolean setAutofilter = true)
            : base(new XLRangeParameters(range.RangeAddress, range.Style ))
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

        #endregion

        private IXLRangeAddress _lastRangeAddress;
        private Dictionary<String, IXLTableField> _fieldNames = null;
        public Dictionary<String, IXLTableField> FieldNames
        {
            get
            {
                if (_fieldNames != null && _lastRangeAddress != null && _lastRangeAddress.Equals(RangeAddress)) return _fieldNames;

                _fieldNames = new Dictionary<String, IXLTableField>();
                _lastRangeAddress = RangeAddress;

                if (ShowHeaderRow)
                {
                    var headersRow = HeadersRow();
                    Int32 cellPos = 0;
                    foreach (var cell in headersRow.Cells())
                    {
                        var name = cell.GetString();
                        if (XLHelper.IsNullOrWhiteSpace(name))
                        {
                            name = "Column" + (cellPos + 1);
                            cell.SetValue(name);
                        }
                        if (_fieldNames.ContainsKey(name))
                            throw new ArgumentException("The header row contains more than one field name '" + name + "'.");

                        _fieldNames.Add(name, new XLTableField(this, name) {Index = cellPos++ });
                    }
                }
                else
                {
                    if (_fieldNames == null) _fieldNames = new Dictionary<String, IXLTableField>();

                    Int32 colCount = ColumnCount();
                    for (Int32 i = 1; i <= colCount; i++)
                    {
                        if (!_fieldNames.Values.Any(f => f.Index == i - 1))
                        {
                            var name = "Column" + i;

                            _fieldNames.Add(name, new XLTableField(this, name) {Index = i - 1 });
                        }
                    }
                }
                return _fieldNames;
            }
        }

        internal void AddFields(IEnumerable<String> fieldNames)
        {
            _fieldNames = new Dictionary<String, IXLTableField>();

            Int32 cellPos = 0;
            foreach(var name in fieldNames)
            {
                _fieldNames.Add(name, new XLTableField(this, name) { Index = cellPos++ });
            }
        }


        internal String RelId { get; set; }

        public IXLTableRange DataRange
        {
            get
            {
                XLRange range;
                //var ws = Worksheet;
                //var tracking = ws.EventTrackingEnabled;
                //ws.EventTrackingEnabled = false;

                if (_showHeaderRow)
                {
                    range = _showTotalsRow
                                ? Range(2, 1,RowCount() - 1,ColumnCount())
                                : Range(2, 1, RowCount(), ColumnCount());
                }
                else
                {
                    range = _showTotalsRow
                                ? Range(1, 1, RowCount() - 1, ColumnCount())
                                : Range(1, 1, RowCount(), ColumnCount());
                }
                //ws.EventTrackingEnabled = tracking;
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
        public Boolean ShowAutoFilter {
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
            if (!ShowHeaderRow) return null;

            var m = FieldNames;
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
                    order = "ASC";
                }

                Int32 co;
                if (!Int32.TryParse(coString, out co))
                    co = Field(coString).Index + 1;

                toSortBy.Append(co);
                toSortBy.Append(" ");
                toSortBy.Append(order);
                toSortBy.Append(",");
            }
            return DataRange.Sort(toSortBy.ToString(0, toSortBy.Length - 1), sortOrder, matchCase, ignoreBlanks);
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

        #endregion



        private void InitializeValues(Boolean setAutofilter)
        {
            ShowRowStripes = true;
            _showHeaderRow = true;
            Theme = XLTableTheme.TableStyleLight9;
            if (setAutofilter)
                InitializeAutoFilter();

            HeadersRow().DataType = XLCellValues.Text;

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
                if (XLHelper.IsNullOrWhiteSpace(((XLCell)c).InnerText))
                    c.Value = GetUniqueName("Column" + co.ToInvariantString());
                _uniqueNames.Add(c.GetString());
                co++;
            }
            Worksheet.Tables.Add(this);
        }


        private String GetUniqueName(String originalName)
        {
            String name = originalName;
            if (_uniqueNames.Contains(name))
            {
                Int32 i = 1;
                name = originalName + i.ToInvariantString();
                while (_uniqueNames.Contains(name))
                {
                    i++;
                    name = originalName + i.ToInvariantString();
                }
            }

            _uniqueNames.Add(name);
            return name;
        }

        public Int32 GetFieldIndex(String name)
        {
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
                        if (XLHelper.IsNullOrWhiteSpace(((XLCell)c).InnerText))
                            c.Value = GetUniqueName("Column" + co.ToInvariantString());
                        _uniqueNames.Add(c.GetString());
                        co++;
                    }

                    headersRow.Clear();
                    RangeAddress.FirstAddress = new XLAddress(Worksheet, RangeAddress.FirstAddress.RowNumber + 1,
                                          RangeAddress.FirstAddress.ColumnNumber,
                                          RangeAddress.FirstAddress.FixedRow,
                                          RangeAddress.FirstAddress.FixedColumn);

                    HeadersRow().DataType = XLCellValues.Text;
                }
                else
                {
                    using(var asRange = Worksheet.Range(
                        RangeAddress.FirstAddress.RowNumber - 1 ,
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

    }
}
