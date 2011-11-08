using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLTable : XLRange, IXLTable
    {
        #region Private fields

        private readonly Dictionary<String, IXLTableField> _fieldNames = new Dictionary<String, IXLTableField>();
        private readonly Dictionary<Int32, IXLTableField> _fields = new Dictionary<Int32, IXLTableField>();
        private string _name;
        internal bool _showTotalsRow;
        internal HashSet<String> _uniqueNames;

        #endregion

        #region Constructor

        public XLTable(XLRange range, Boolean addToTables)
            : base(range.RangeParameters)
        {
            InitializeValues();

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

        public XLTable(XLRange range, String name, Boolean addToTables)
            : base(range.RangeParameters)
        {
            InitializeValues();

            Name = name;
            AddToTables(range, addToTables);
        }

        #endregion

        public String RelId { get; set; }

        public IXLRange DataRange
        {
            get
            {
                return _showTotalsRow ? Range(2, 1, RowCount() - 1, ColumnCount()) : Range(2, 1, RowCount(), ColumnCount());
            }
        }

        #region IXLTable Members

        public Boolean EmphasizeFirstColumn { get; set; }
        public Boolean EmphasizeLastColumn { get; set; }
        public Boolean ShowRowStripes { get; set; }
        public Boolean ShowColumnStripes { get; set; }
        public Boolean ShowAutoFilter { get; set; }
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
            }
        }

        public IXLRangeRow HeadersRow()
        {
            return new XLTableRow(this, base.FirstRow());
        }

        public IXLRangeRow TotalsRow()
        {
            if (ShowTotalsRow)
                return new XLTableRow(this, (XLRangeRow)base.LastRow());

            throw new InvalidOperationException("Cannot access TotalsRow if ShowTotals property is false");
        }

        public new IXLTableRow FirstRow()
        {
            return Row(1);
        }

        public new IXLTableRow FirstRowUsed()
        {
            return new XLTableRow(this, (XLRangeRow)(DataRange.FirstRowUsed()));
        }

        public new IXLTableRow LastRow()
        {
            return ShowTotalsRow ? new XLTableRow(this, base.Row(RowCount() - 1)) : new XLTableRow(this, base.Row(RowCount()));
        }

        public new IXLTableRow LastRowUsed()
        {
            return new XLTableRow(this, (XLRangeRow)(DataRange.LastRowUsed()));
        }

        IXLTableRow IXLTable.Row(int row)
        {
            return Row(row);
        }

        public new IXLTableRows Rows()
        {
            var retVal = new XLTableRows(Worksheet.Style);
            foreach (int r in Enumerable.Range(1, DataRange.RowCount()))
                retVal.Add(Row(r));
            return retVal;
        }

        public new IXLTableRows Rows(int firstRow, int lastRow)
        {
            var retVal = new XLTableRows(Worksheet.Style);

            for (int ro = firstRow; ro <= lastRow; ro++)
                retVal.Add(Row(ro));
            return retVal;
        }

        public new IXLTableRows Rows(string rows)
        {
            var retVal = new XLTableRows(Worksheet.Style);
            var rowPairs = rows.Split(',');
            foreach (string tPair in rowPairs.Select(pair => pair.Trim()))
            {
                String firstRow;
                String lastRow;
                if (tPair.Contains(':') || tPair.Contains('-'))
                {
                    string[] rowRange = ExcelHelper.SplitRange(tPair);

                    firstRow = rowRange[0];
                    lastRow = rowRange[1];
                }
                else
                {
                    firstRow = tPair;
                    lastRow = tPair;
                }
                foreach (IXLTableRow row in Rows(Int32.Parse(firstRow), Int32.Parse(lastRow)))
                    retVal.Add(row);
            }
            return retVal;
        }

        public new IXLRangeColumn Column(Int32 column)
        {
            return DataRange.Column(column);
        }

        public new IXLRangeColumn Column(String column)
        {
            if (ExcelHelper.IsValidColumn(column))
            {
                Int32 coNum = ExcelHelper.GetColumnNumberFromLetter(column);
                return coNum > ColumnCount() ? DataRange.Column(GetFieldIndex(column) + 1) : DataRange.Column(coNum);
            }

            return DataRange.Column(GetFieldIndex(column) + 1);
        }

        public new IXLRangeColumns Columns()
        {
            return DataRange.Columns();
        }

        public new IXLRangeColumns Columns(Int32 firstColumn, Int32 lastColumn)
        {
            return DataRange.Columns(firstColumn, lastColumn);
        }

        public new IXLRangeColumns Columns(String firstColumn, String lastColumn)
        {
            return DataRange.Columns(firstColumn, lastColumn);
        }

        public new IXLRangeColumns Columns(String columns)
        {
            return DataRange.Columns(columns);
        }

        IXLCell IXLTable.Cell(int row, int column)
        {
            return Cell(row, column);
        }

        IXLCell IXLTable.Cell(string cellAddressInRange)
        {
            return Cell(cellAddressInRange);
        }

        IXLCell IXLTable.Cell(int row, string column)
        {
            return Cell(row, column);
        }

        IXLCell IXLTable.Cell(IXLAddress cellAddressInRange)
        {
            return Cell(cellAddressInRange);
        }

        IXLRange IXLTable.Range(IXLRangeAddress rangeAddress)
        {
            return Range(rangeAddress);
        }

        IXLRange IXLTable.Range(string rangeAddress)
        {
            return Range(rangeAddress);
        }

        IXLRange IXLTable.Range(IXLCell firstCell, IXLCell lastCell)
        {
            return Range(firstCell, lastCell);
        }

        IXLRange IXLTable.Range(string firstCellAddress, string lastCellAddress)
        {
            return Range(firstCellAddress, lastCellAddress);
        }

        IXLRange IXLTable.Range(IXLAddress firstCellAddress, IXLAddress lastCellAddress)
        {
            return Range(firstCellAddress, lastCellAddress);
        }

        IXLRange IXLTable.Range(int firstCellRow, int firstCellColumn, int lastCellRow, int lastCellColumn)
        {
            return Range(firstCellRow, firstCellColumn, lastCellRow, lastCellColumn);
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


        IXLRangeColumn IXLTable.FirstColumn()
        {
            return FirstColumn();
        }

        IXLRangeColumn IXLTable.FirstColumnUsed()
        {
            return FirstColumnUsed();
        }

        IXLRangeColumn IXLTable.LastColumnUsed()
        {
            return LastColumnUsed();
        }

        public new IXLRange Sort(String columnsToSortBy, XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true)
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

        #endregion

        public new XLTableRow Row(int row)
        {
            if (row <= 0 || row > ExcelHelper.MaxRowNumber)
            {
                throw new IndexOutOfRangeException(String.Format("Row number must be between 1 and {0}",
                                                                 ExcelHelper.MaxRowNumber));
            }

            return new XLTableRow(this, base.Row(row + 1));
        }

        private void InitializeValues()
        {
            ShowRowStripes = true;
            ShowAutoFilter = true;
            Theme = XLTableTheme.TableStyleLight9;
            AutoFilter = new XLAutoFilter();
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
            if (_fieldNames.ContainsKey(name))
                return _fieldNames[name].Index;

            var headersRow = HeadersRow();
            Int32 cellCount = headersRow.CellCount();
            for (Int32 cellPos = 1; cellPos <= cellCount; cellPos++)
            {
                if (!headersRow.Cell(cellPos).GetString().Equals(name)) continue;

                if (_fieldNames.ContainsKey(name))
                {
                    throw new ArgumentException("The header row contains more than one field name '" + name +
                                                "'.");
                }
                _fieldNames.Add(name, Field(cellPos - 1));
            }
            if (_fieldNames.ContainsKey(name))
                return _fieldNames[name].Index;

            throw new ArgumentOutOfRangeException("The header row doesn't contain field name '" + name + "'.");
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
        public XLAutoFilter AutoFilter { get; private set; }
    }
}