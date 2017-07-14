using System;

namespace ClosedXML.Excel
{
    internal class XLTableField: IXLTableField
    {
        private XLTable table;
        public XLTableField(XLTable table, String name)
        {
            this.table = table;
            this.name = name;
        }

        public Int32 Index { get; internal set; }

        private String name;

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

        internal String totalsRowLabel;
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

        internal XLTotalsRowFunction totalsRowFunction;
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
    }
}
