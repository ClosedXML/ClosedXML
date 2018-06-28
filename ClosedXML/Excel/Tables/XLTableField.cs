using System;
using System.Collections.Generic;
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

        private IXLRangeColumn _column;
        private Int32 index;
        private String name;

        public XLTableField(XLTable table, String name)
        {
            this.table = table;
            this.name = name;
        }

        public IXLRangeColumn Column
        {
            get
            {
                if (_column == null)
                {
                    _column = table.AsRange().Column(this.Index + 1);
                }
                return _column;
            }
            internal set
            {
                _column = value;
            }
        }

        public IXLCells DataCells
        {
            get
            {
                return Column.Cells(c =>
                {
                    if (table.ShowHeaderRow && c == HeaderCell)
                        return false;
                    if (table.ShowTotalsRow && c == TotalsCell)
                        return false;
                    return true;
                });
            }
        }

        public IXLCell HeaderCell
        {
            get
            {
                if (!table.ShowHeaderRow)
                    return null;

                return Column.FirstCell();
            }
        }

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
                if (name == value) return;

                if (table.ShowHeaderRow)
                    (table.HeadersRow(false).Cell(Index + 1) as XLCell).SetValue(value, false);

                table.RenameField(name, value);
                name = value;
            }
        }

        public IXLTable Table { get { return table; } }

        public IXLCell TotalsCell
        {
            get
            {
                if (!table.ShowTotalsRow)
                    return null;

                return Column.LastCell();
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
                UpdateTableFieldTotalsRowFormula();
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
            }

            fields.Where(f => f.Index > this.Index).ForEach(f => f.Index--);
            table.FieldNames.Remove(this.Name);
        }

        public bool IsConsistentDataType()
        {
            var dataTypes = this.Column
                .Cells()
                .Skip(this.table.ShowHeaderRow ? 1 : 0)
                .Select(c => c.DataType);

            if (this.table.ShowTotalsRow)
                dataTypes = dataTypes.Take(dataTypes.Count() - 1);

            var distinctDataTypes = dataTypes
                .GroupBy(dt => dt)
                .Select(g => new { Key = g.Key, Count = g.Count() });

            return distinctDataTypes.Count() == 1;
        }

        public Boolean IsConsistentFormula()
        {
            var formulas = this.Column
                .Cells()
                .Skip(this.table.ShowHeaderRow ? 1 : 0)
                .Select(c => c.FormulaR1C1);

            if (this.table.ShowTotalsRow)
                formulas = formulas.Take(formulas.Count() - 1);

            var distinctFormulas = formulas
                .GroupBy(f => f)
                .Select(g => new { Key = g.Key, Count = g.Count() });

            return distinctFormulas.Count() == 1;
        }

        public bool IsConsistentStyle()
        {
            var styles = this.Column
                .Cells()
                .Skip(this.table.ShowHeaderRow ? 1 : 0)
                .OfType<XLCell>()
                .Select(c => c.StyleValue);

            if (this.table.ShowTotalsRow)
                styles = styles.Take(styles.Count() - 1);

            var distinctStyles = styles
                .Distinct();

            return distinctStyles.Count() == 1;
        }

        private static IEnumerable<String> QuotedTableFieldCharacters = new[] { "'", "#" };

        internal void UpdateTableFieldTotalsRowFormula()
        {
            if (TotalsRowFunction != XLTotalsRowFunction.None && TotalsRowFunction != XLTotalsRowFunction.Custom)
            {
                var cell = table.TotalsRow().Cell(Index + 1);
                var formulaCode = String.Empty;
                switch (TotalsRowFunction)
                {
                    case XLTotalsRowFunction.Sum: formulaCode = "109"; break;
                    case XLTotalsRowFunction.Minimum: formulaCode = "105"; break;
                    case XLTotalsRowFunction.Maximum: formulaCode = "104"; break;
                    case XLTotalsRowFunction.Average: formulaCode = "101"; break;
                    case XLTotalsRowFunction.Count: formulaCode = "103"; break;
                    case XLTotalsRowFunction.CountNumbers: formulaCode = "102"; break;
                    case XLTotalsRowFunction.StandardDeviation: formulaCode = "107"; break;
                    case XLTotalsRowFunction.Variance: formulaCode = "110"; break;
                }

                var modifiedName = Name;
                QuotedTableFieldCharacters.ForEach(c => modifiedName = modifiedName.Replace(c, "'" + c));

                if (modifiedName.StartsWith(" ") || modifiedName.EndsWith(" "))
                {
                    modifiedName = "[" + modifiedName + "]";
                }

                var prependTableName = modifiedName.Contains(" ");

                cell.FormulaA1 = $"SUBTOTAL({formulaCode},{(prependTableName ? table.Name : string.Empty)}[{modifiedName}])";
                var lastCell = table.LastRow().Cell(Index + 1);
                if (lastCell.DataType != XLDataType.Text)
                {
                    cell.DataType = lastCell.DataType;
                    cell.Style.NumberFormat = lastCell.Style.NumberFormat;
                }
            }
        }
    }
}
