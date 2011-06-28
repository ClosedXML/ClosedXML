using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLTable : XLRange, IXLTable
    {
        #region Private fields
        private string m_name;
        internal bool m_showTotalsRow;
        internal Dictionary<Int32, IXLTableField> m_fields = new Dictionary<Int32, IXLTableField>();
        private readonly Dictionary<String, IXLTableField> m_fieldNames = new Dictionary<String, IXLTableField>();
        internal HashSet<String> m_uniqueNames;
        #endregion
        #region Constructor
        public XLTable(XLRange range, Boolean addToTables)
                : base(range.RangeParameters)
        {
            InitializeValues();

            Int32 id = 1;
            while (true)
            {
                var tableName = String.Format("Table{0}", id);
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
        public Boolean EmphasizeFirstColumn { get; set; }
        public Boolean EmphasizeLastColumn { get; set; }
        public Boolean ShowRowStripes { get; set; }
        public Boolean ShowColumnStripes { get; set; }
        public Boolean ShowAutoFilter { get; set; }
        public XLTableTheme Theme { get; set; }

        public String Name
        {
            get { return m_name; }
            set
            {
                if (Worksheet.Tables.Any(t => t.Name == value))
                {
                    throw new ArgumentException(String.Format("This worksheet already contains a table named '{0}'", value));
                }

                m_name = value;
            }
        }

        public Boolean ShowTotalsRow
        {
            get { return m_showTotalsRow; }
            set
            {
                if (value && !m_showTotalsRow)
                {
                    InsertRowsBelow(1);
                }
                else if (!value && m_showTotalsRow)
                {
                    TotalsRow().Delete();
                }

                m_showTotalsRow = value;
            }
        }

        public IXLRange DataRange
        {
            get
            {
                if (m_showTotalsRow)
                {
                    return base.Range(2, 1, RowCount() - 1, ColumnCount());
                }
                else
                {
                    return base.Range(2, 1, RowCount(), ColumnCount());
                }
            }
        }

        private void InitializeValues()
        {
            ShowRowStripes = true;
            ShowAutoFilter = true;
            Theme = XLTableTheme.TableStyleLight9;
        }

        private void AddToTables(XLRange range, Boolean addToTables)
        {
            if (addToTables)
            {
                m_uniqueNames = new HashSet<string>();
                Int32 co = 1;
                foreach (var c in range.Row(1).Cells())
                {
                    if (StringExtensions.IsNullOrWhiteSpace(((XLCell) c).InnerText))
                    {
                        c.Value = GetUniqueName("Column" + co.ToStringLookup());
                    }
                    m_uniqueNames.Add(c.GetString());
                    co++;
                }
                Worksheet.Tables.Add(this);
            }
        }

        private String GetUniqueName(String originalName)
        {
            String name = originalName;
            if (m_uniqueNames.Contains(name))
            {
                Int32 i = 1;
                name = originalName + i.ToStringLookup();
                while (m_uniqueNames.Contains(name))
                {
                    i++;
                    name = originalName + i.ToStringLookup();
                }
            }

            m_uniqueNames.Add(name);
            return name;
        }

        public IXLRangeRow HeadersRow()
        {
            return new XLTableRow(this, (XLRangeRow) base.FirstRow());
        }

        public IXLRangeRow TotalsRow()
        {
            if (ShowTotalsRow)
            {
                return new XLTableRow(this, (XLRangeRow) base.LastRow());
            }
            else
            {
                throw new InvalidOperationException("Cannot access TotalsRow if ShowTotals property is false");
            }
        }

        public new IXLTableRow FirstRow()
        {
            return Row(1);
        }

        public new IXLTableRow FirstRowUsed()
        {
            return new XLTableRow(this, (XLRangeRow) (DataRange.FirstRowUsed()));
        }

        public new IXLTableRow LastRow()
        {
            if (ShowTotalsRow)
            {
                return new XLTableRow(this, (XLRangeRow) base.Row(RowCount() - 1));
            }
            else
            {
                return new XLTableRow(this, (XLRangeRow) base.Row(RowCount()));
            }
        }

        public new IXLTableRow LastRowUsed()
        {
            return new XLTableRow(this, (XLRangeRow) (DataRange.LastRowUsed()));
        }

        public new IXLTableRow Row(int row)
        {
            if (row <= 0 || row > ExcelHelper.MaxRowNumber)
                throw new IndexOutOfRangeException(String.Format("Row number must be between 1 and {0}", ExcelHelper.MaxRowNumber));

            return new XLTableRow(this, (XLRangeRow) base.Row(row + 1));
        }

        public new IXLTableRows Rows()
        {
            var retVal = new XLTableRows(Worksheet.Style);
            foreach (var r in Enumerable.Range(1, DataRange.RowCount()))
            {
                retVal.Add(Row(r));
            }
            return retVal;
        }

        public new IXLTableRows Rows(int firstRow, int lastRow)
        {
            var retVal = new XLTableRows(Worksheet.Style);

            for (var ro = firstRow; ro <= lastRow; ro++)
            {
                retVal.Add(Row(ro));
            }
            return retVal;
        }

        public new IXLTableRows Rows(string rows)
        {
            var retVal = new XLTableRows(Worksheet.Style);
            var rowPairs = rows.Split(',');
            foreach (var pair in rowPairs)
            {
                var tPair = pair.Trim();
                String firstRow;
                String lastRow;
                if (tPair.Contains(':') || tPair.Contains('-'))
                {
                    if (tPair.Contains('-'))
                    {
                        tPair = tPair.Replace('-', ':');
                    }

                    var rowRange = tPair.Split(':');
                    firstRow = rowRange[0];
                    lastRow = rowRange[1];
                }
                else
                {
                    firstRow = tPair;
                    lastRow = tPair;
                }
                foreach (var row in Rows(Int32.Parse(firstRow), Int32.Parse(lastRow)))
                {
                    retVal.Add(row);
                }
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
                if (coNum > ColumnCount())
                {
                    return DataRange.Column(GetFieldIndex(column) + 1);
                }
                else
                {
                    return DataRange.Column(coNum);
                }
            }
            else
            {
                return DataRange.Column(GetFieldIndex(column) + 1);
            }
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
            if (!m_fields.ContainsKey(fieldIndex))
            {
                if (fieldIndex >= HeadersRow().CellCount())
                {
                    throw new ArgumentOutOfRangeException();
                }

                var newField = new XLTableField(this) {Index = fieldIndex, Name = HeadersRow().Cell(fieldIndex + 1).GetString()};
                m_fields.Add(fieldIndex, newField);
            }

            return m_fields[fieldIndex];
        }

        public Int32 GetFieldIndex(String name)
        {
            if (m_fieldNames.ContainsKey(name))
            {
                return m_fieldNames[name].Index;
            }
            else
            {
                var headersRow = HeadersRow();
                Int32 cellCount = headersRow.CellCount();
                for (Int32 cellPos = 1; cellPos <= cellCount; cellPos++)
                {
                    if (headersRow.Cell(cellPos).GetString().Equals(name))
                    {
                        if (m_fieldNames.ContainsKey(name))
                        {
                            throw new ArgumentException("The header row contains more than one field name '" + name + "'.");
                        }
                        else
                        {
                            m_fieldNames.Add(name, Field(cellPos - 1));
                        }
                    }
                }
                if (m_fieldNames.ContainsKey(name))
                {
                    return m_fieldNames[name].Index;
                }
                else
                {
                    throw new ArgumentOutOfRangeException("The header row doesn't contain field name '" + name + "'.");
                }
            }
        }

        public new IXLRange Sort(String elementsToSortBy)
        {
            StringBuilder toSortBy = new StringBuilder();
            foreach (String coPair in elementsToSortBy.Split(','))
            {
                String coPairTrimmed = coPair.Trim();
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
                {
                    co = Field(coString).Index + 1;
                }

                toSortBy.Append(co);
                toSortBy.Append(" ");
                toSortBy.Append(order);
                toSortBy.Append(",");
            }
            return DataRange.Sort(toSortBy.ToString(0, toSortBy.Length - 1));
        }

        public IXLTable SetEmphasizeFirstColumn() { EmphasizeFirstColumn = true; return this; }	public IXLTable SetEmphasizeFirstColumn(Boolean value) { EmphasizeFirstColumn = value; return this; }
        public IXLTable SetEmphasizeLastColumn() { EmphasizeLastColumn = true; return this; }	public IXLTable SetEmphasizeLastColumn(Boolean value) { EmphasizeLastColumn = value; return this; }
        public IXLTable SetShowRowStripes() { ShowRowStripes = true; return this; }	public IXLTable SetShowRowStripes(Boolean value) { ShowRowStripes = value; return this; }
        public IXLTable SetShowColumnStripes() { ShowColumnStripes = true; return this; }	public IXLTable SetShowColumnStripes(Boolean value) { ShowColumnStripes = value; return this; }
        public IXLTable SetShowTotalsRow() { ShowTotalsRow = true; return this; }	public IXLTable SetShowTotalsRow(Boolean value) { ShowTotalsRow = value; return this; }
        public IXLTable SetShowAutoFilter() { ShowAutoFilter = true; return this; }	public IXLTable SetShowAutoFilter(Boolean value) { ShowAutoFilter = value; return this; }



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
    }
}