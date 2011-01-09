using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLTable : XLRange, IXLTable
    {
        public String RelId { get; set; }
        public String Name { get; set; }
        public Boolean EmphasizeFirstColumn { get; set; }
        public Boolean EmphasizeLastColumn { get; set; }
        public Boolean ShowRowStripes { get; set; }
        public Boolean ShowColumnStripes { get; set; }
        public Boolean ShowAutoFilter { get; set; }
        public XLTableTheme Theme { get; set; }

        internal Boolean showTotalsRow;
        public Boolean ShowTotalsRow
        {
            get { return showTotalsRow; }
            set 
            {
                if (value && !showTotalsRow)
                    this.InsertRowsBelow(1);
                else if (!value && showTotalsRow)
                    this.TotalsRow().Delete();

                showTotalsRow = value; 
            }
        }

        public IXLRange DataRange
        {
            get
            { 
                if (showTotalsRow)
                    return base.Range(2,1, RowCount() - 1, ColumnCount());
                else
                    return base.Range(2, 1, RowCount(), ColumnCount());
            }
        }
        public XLTable(XLRange range, Boolean addToTables)
            : base(range.RangeParameters)
        {
            InitializeValues();

            Int32 id = 1;
            while (true)
            {
                String tableName = String.Format("Table{0}", id);
                if (!Worksheet.Tables.Where(t=>t.Name == tableName).Any())
                {
                    Name = tableName;
                    if (addToTables)
                        Worksheet.Tables.Add(this);
                    break;
                }
                id++;
            }
        }

        private void InitializeValues()
        {
            ShowRowStripes = true;
            ShowAutoFilter = true;
            Theme = XLTableTheme.TableStyleLight9;
        }

        public XLTable(XLRange range, String name, Boolean addToTables)
            : base(range.RangeParameters)
        {
            InitializeValues();

            this.Name = name;
            if (addToTables)
                Worksheet.Tables.Add(this);
        }


        public IXLRangeRow HeadersRow()
        {
            return new XLTableRow(this, (XLRangeRow)base.FirstRow());
        }

        public IXLRangeRow TotalsRow()
        {
            if (ShowTotalsRow)
                return new XLTableRow(this, (XLRangeRow)base.LastRow());
            else
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
            if (ShowTotalsRow)
                return new XLTableRow(this, (XLRangeRow)base.Row(RowCount() - 1));
            else
                return new XLTableRow(this, (XLRangeRow)base.Row(RowCount()));
        }

        public new IXLTableRow LastRowUsed()
        {
            return new XLTableRow(this, (XLRangeRow)(DataRange.LastRowUsed()));
        }

        public new IXLTableRow Row(int row)
        {
            return new XLTableRow(this, (XLRangeRow)base.Row(row + 1));
        }

        public new IXLTableRows Rows()
        {
            var retVal = new XLTableRows(Worksheet);
            foreach (var r in Enumerable.Range(1, DataRange.RowCount()))
            {
                retVal.Add(this.Row(r));
            }
            return retVal;
        }

        public new IXLTableRows Rows(int firstRow, int lastRow)
        {
            var retVal = new XLTableRows(Worksheet);

            for (var ro = firstRow; ro <= lastRow; ro++)
            {
                retVal.Add(this.Row(ro));
            }
            return retVal;
        }

        public new IXLTableRows Rows(string rows)
        {
            var retVal = new XLTableRows(Worksheet);
            var rowPairs = rows.Split(',');
            foreach (var pair in rowPairs)
            {
                var tPair = pair.Trim();
                String firstRow;
                String lastRow;
                if (tPair.Contains(':') || tPair.Contains('-'))
                {
                    if (tPair.Contains('-'))
                        tPair = tPair.Replace('-', ':');

                    var rowRange = tPair.Split(':');
                    firstRow = rowRange[0];
                    lastRow = rowRange[1];
                }
                else
                {
                    firstRow = tPair;
                    lastRow = tPair;
                }
                foreach (var row in this.Rows(Int32.Parse(firstRow), Int32.Parse(lastRow)))
                {
                    retVal.Add(row);
                }
            }
            return retVal;
        }

        public IXLTableField Field(String fieldName)
        {
            return Field(GetFieldIndex(fieldName));
        }

        private Dictionary<Int32, IXLTableField> fields = new Dictionary<Int32, IXLTableField>();
        public IXLTableField Field(Int32 fieldIndex)
        {
            if (!fields.ContainsKey(fieldIndex))
            {
                if (fieldIndex >= HeadersRow().CellCount())
                    throw new ArgumentOutOfRangeException();

                var newField = new XLTableField(this) { Index = fieldIndex, Name = HeadersRow().Cell(fieldIndex + 1).GetString() };
                fields.Add(fieldIndex, newField);
            }

            return fields[fieldIndex];
        }

        private Dictionary<String, IXLTableField> fieldNames = new Dictionary<String, IXLTableField>();
        public  Int32 GetFieldIndex(String name)
        {
            if (fieldNames.ContainsKey(name))
            {
                return fieldNames[name].Index;
            }
            else
            {
                var headersRow = HeadersRow();
                Int32 cellCount = headersRow.CellCount();
                for (Int32 cellPos = 1; cellPos <= cellCount; cellPos++)
                {
                    if (headersRow.Cell(cellPos).GetString().Equals(name))
                    {
                        if (fieldNames.ContainsKey(name))
                        {
                            throw new ArgumentException("The header row contains more than one field name '" + name + "'.");
                        }
                        else
                        {
                            fieldNames.Add(name, Field(cellPos - 1));
                        }
                    }
                }
                if (fieldNames.ContainsKey(name))
                {
                    return fieldNames[name].Index;
                }
                else
                {
                    throw new ArgumentOutOfRangeException("The header row doesn't contain field name '" + name + "'.");
                }
            }
        }
    }
}
