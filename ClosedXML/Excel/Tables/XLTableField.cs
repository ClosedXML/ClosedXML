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

        public IXLRangeColumn Column
        {
            get { return table.Column(this.Index); }
        }

        public Int32 Index { get; internal set; }

        public String Name
        {
            get
            {
                return name;
            }
            set
            {
                if (table.ShowHeaderRow)
                    table.HeadersRow().Cell(Index + 1).SetValue(value);

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
                if (value != XLTotalsRowFunction.None && value != XLTotalsRowFunction.Custom)
                {
                    var cell = table.TotalsRow().Cell(Index + 1);
                    String formula = String.Empty;
                    switch (value)
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
                totalsRowFunction = value;
            }
        }

        public String TotalsRowLabel
        {
            get { return totalsRowLabel; }
            set
            {
                totalsRowFunction = XLTotalsRowFunction.None;
                table.TotalsRow().Cell(Index + 1).SetValue(value);
                totalsRowLabel = value;
            }
        }

        public void Delete()
        {
            Delete(true);
        }

        internal void Delete(Boolean deleteUnderlyingRangeColumn)
        {
            var fields = table.Fields.Cast<XLTableField>();
            fields.Where(f => f.Index > this.Index).ForEach(f => f.Index--);
            table.FieldNames.Remove(this.Name);

            if (deleteUnderlyingRangeColumn)
                (this.Column as XLRangeColumn).Delete(false);
        }
    }
}
