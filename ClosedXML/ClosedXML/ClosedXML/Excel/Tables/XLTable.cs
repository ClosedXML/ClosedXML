using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLTable : XLRange, IXLTable
    {
        #region Private fields

        public readonly Dictionary<String, IXLTableField> FieldNames = new Dictionary<String, IXLTableField>();
        private readonly Dictionary<Int32, IXLTableField> _fields = new Dictionary<Int32, IXLTableField>();
        private string _name;
        internal bool _showTotalsRow;
        internal HashSet<String> _uniqueNames;

        #endregion

        #region Constructor

        public XLTable(XLRange range, Boolean addToTables, Boolean setAutofilter = true)
            : base(range.RangeParameters)
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
            : base(range.RangeParameters)
        {
            InitializeValues(setAutofilter);

            Name = name;
            AddToTables(range, addToTables);
        }

        #endregion

        public String RelId { get; set; }

        public IXLTableRange DataRange
        {
            get
            {
                var range = _showTotalsRow
                           ? Range(2, 1, RowCount() - 1, ColumnCount())
                           : Range(2, 1, RowCount(), ColumnCount());
                return new XLTableRange(range, this);
            }
        }

        public XLAutoFilter AutoFilter { get; private set; }

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
            return ShowHeaderRow ? FirstRow() : null;
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
            if (!_fields.ContainsKey(fieldIndex))
            {
                if (fieldIndex >= HeadersRow().CellCount())
                    throw new ArgumentOutOfRangeException();

                var newField = new XLTableField(this)
                                   {Index = fieldIndex, Name = HeadersRow().Cell(fieldIndex + 1).GetString()};
                _fields.Add(fieldIndex, newField);
            }

            return _fields[fieldIndex];
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
        }

        public void InitializeAutoFilter()
        {
            AutoFilter = new XLAutoFilter { Range = AsRange() };
            ShowAutoFilter = true;
        }

        private void AddToTables(XLRange range, Boolean addToTables)
        {
            if (!addToTables) return;

            _uniqueNames = new HashSet<string>();
            Int32 co = 1;
            foreach (IXLCell c in range.Row(1).Cells())
            {
                if (StringExtensions.IsNullOrWhiteSpace(((XLCell)c).InnerText))
                    c.Value = GetUniqueName("Column" + co.ToStringLookup());
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
                name = originalName + i.ToStringLookup();
                while (_uniqueNames.Contains(name))
                {
                    i++;
                    name = originalName + i.ToStringLookup();
                }
            }

            _uniqueNames.Add(name);
            return name;
        }

        public Int32 GetFieldIndex(String name)
        {
            if (FieldNames.ContainsKey(name))
                return FieldNames[name].Index;

            var headersRow = HeadersRow();
            Int32 cellCount = headersRow.CellCount();
            for (Int32 cellPos = 1; cellPos <= cellCount; cellPos++)
            {
                if (!headersRow.Cell(cellPos).GetString().Equals(name)) continue;

                if (FieldNames.ContainsKey(name))
                {
                    throw new ArgumentException("The header row contains more than one field name '" + name +
                                                "'.");
                }
                FieldNames.Add(name, Field(cellPos - 1));
            }
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
                        if (StringExtensions.IsNullOrWhiteSpace(((XLCell)c).InnerText))
                            c.Value = GetUniqueName("Column" + co.ToStringLookup());
                        _uniqueNames.Add(c.GetString());
                        co++;
                    }
                    _uniqueNames.ForEach(n=>AddField(n));
                    headersRow.Clear();

                }
                else
                {
                    using(var asRange = AsRange())
                        using (var firstRow = asRange.FirstRow())
                            {
                                IXLRangeRow rangeRow;
                                if (firstRow.IsEmpty(true))
                                    rangeRow = firstRow;
                                else
                                {
                                    rangeRow = firstRow.InsertRowsBelow(1).First();
                                    
                                    RangeAddress.FirstAddress = new XLAddress(Worksheet, RangeAddress.FirstAddress.RowNumber + 1,
                                                                              RangeAddress.FirstAddress.ColumnNumber,
                                                                              RangeAddress.FirstAddress.FixedRow,
                                                                              RangeAddress.FirstAddress.FixedColumn);

                                    RangeAddress.LastAddress = new XLAddress(Worksheet, RangeAddress.LastAddress.RowNumber,
                                                                             RangeAddress.LastAddress.ColumnNumber,
                                                                             RangeAddress.LastAddress.FixedRow,
                                                                             RangeAddress.LastAddress.FixedColumn);
                                }

                                Int32 co = 1;
                                foreach (var name in FieldNames.Keys)
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

        public XLTable AddField(String name)
        {
            var field = new XLTableField(this) {Index = _fields.Count, Name = name};
            if (!_fields.ContainsKey(_fields.Count))
                _fields.Add(_fields.Count, field);

            if (!FieldNames.ContainsKey(name))
                FieldNames.Add(name, field);

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