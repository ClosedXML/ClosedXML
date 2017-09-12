using System;
using System.Diagnostics;
using System.Linq;

namespace ClosedXML.Excel
{
    [DebuggerDisplay("{Name}")]
    internal class XLTableField : IXLTableField
    {
        internal XLTotalsRowFunction totalsRowFunction;
        internal String totalsRowLabel;
        private readonly XLTable table;

        private String name;

        public XLTableField(XLTable table, String name)
        {
            this.table = table;
            this.name = name;
        }

        private IXLRangeColumn _column;

        public IXLRangeColumn Column
        {
            get { return _column ?? (_column = table.HeadersRow(false).Cell(this.Index + 1).AsRange().Columns().Single()); }
        }

        private Int32 index;

        public Int32 Index
        {
            get { return index; }
            internal set
            {
                if (index == value) return;
                index = value;
                _column = null;
            }
        }

        public String Name
        {
            get
            {
                return name;
            }
            set
            {
                if (table.ShowHeaderRow)
                    (table.HeadersRow(false).Cell(Index + 1) as XLCell).SetValue(value, false);

                table.RenameField(name, value);
                name = value;
            }
        }

        public String TotalsRowFormulaA1
        {
            get { return table.TotalsRow().Cell(Index + 1).FormulaA1; }
            set
            {
                totalsRowFunction = XLTotalsRowFunction.Custom;
                table.TotalsRow().Cell(Index + 1).FormulaA1 = value;
            }
        }

        public String TotalsRowFormulaR1C1
        {
            get { return table.TotalsRow().Cell(Index + 1).FormulaR1C1; }
            set
            {
                totalsRowFunction = XLTotalsRowFunction.Custom;
                table.TotalsRow().Cell(Index + 1).FormulaR1C1 = value;
            }
        }

        public XLTotalsRowFunction TotalsRowFunction
        {
            get { return totalsRowFunction; }
            set
            {
                totalsRowFunction = value;
                UpdateUnderlyingCellFormula();
            }
        }

        internal void UpdateUnderlyingCellFormula()
        {
            if (TotalsRowFunction != XLTotalsRowFunction.None && TotalsRowFunction != XLTotalsRowFunction.Custom)
            {
                var cell = table.TotalsRow().Cell(Index + 1);
                String formula = String.Empty;
                switch (TotalsRowFunction)
                {
                    case XLTotalsRowFunction.Sum: formula = "109"; break;
                    case XLTotalsRowFunction.Minimum: formula = "105"; break;
                    case XLTotalsRowFunction.Maximum: formula = "104"; break;
                    case XLTotalsRowFunction.Average: formula = "101"; break;
                    case XLTotalsRowFunction.Count: formula = "103"; break;
                    case XLTotalsRowFunction.CountNumbers: formula = "102"; break;
                    case XLTotalsRowFunction.StandardDeviation: formula = "107"; break;
                    case XLTotalsRowFunction.Variance: formula = "110"; break;
                }

                cell.FormulaA1 = "SUBTOTAL(" + formula + ",[" + Name + "])";
                var lastCell = table.LastRow().Cell(Index + 1);
                if (lastCell.DataType != XLCellValues.Text)
                {
                    cell.DataType = lastCell.DataType;
                    cell.Style.NumberFormat = lastCell.Style.NumberFormat;
                }
            }
        }

        public String TotalsRowLabel
        {
            get { return totalsRowLabel; }
            set
            {
                totalsRowFunction = XLTotalsRowFunction.None;
                (table.TotalsRow().Cell(Index + 1) as XLCell).SetValue(value, false);
                totalsRowLabel = value;
            }
        }

        public void Delete()
        {
            Delete(true);
        }

        internal void Delete(Boolean deleteUnderlyingRangeColumn)
        {
            var fields = table.Fields.Cast<XLTableField>().ToArray();

            if (deleteUnderlyingRangeColumn)
            {
                table.AsRange().ColumnQuick(this.Index + 1).Delete();
                // (this.Column as XLRangeColumn).Delete(false);
            }

            fields.Where(f => f.Index > this.Index).ForEach(f => f.Index--);
            table.FieldNames.Remove(this.Name);
        }
    }
}