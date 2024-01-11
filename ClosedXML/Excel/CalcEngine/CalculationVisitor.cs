using System;
using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using ClosedXML.Parser;

namespace ClosedXML.Excel.CalcEngine
{
    internal class CalculationVisitor : IFormulaVisitor<CalcContext, AnyValue>
    {
        private readonly FunctionRegistry _functions;
        private readonly ArrayPool<AnyValue> _argsPool;

        public CalculationVisitor(FunctionRegistry functions)
        {
            _functions = functions;
            _argsPool = ArrayPool<AnyValue>.Create(XLConstants.MaxFunctionArguments, 100);
        }

        public AnyValue Visit(CalcContext context, ScalarNode node)
        {
            return node.Value.ToAnyValue();
        }

        public AnyValue Visit(CalcContext context, ArrayNode node)
        {
            return node.Value;
        }

        public AnyValue Visit(CalcContext context, UnaryNode node)
        {
            var arg = node.Expression.Accept(context, this);

            return node.Operation switch
            {
                UnaryOp.Add => arg.UnaryPlus(),
                UnaryOp.Subtract => arg.UnaryMinus(context),
                UnaryOp.Percentage => arg.UnaryPercent(context),
                UnaryOp.SpillRange => throw new NotImplementedException("Evaluation of spill range operator is not implemented."),
                UnaryOp.ImplicitIntersection => throw new NotImplementedException("Excel 2016 implicit intersection is different from @ intersection of E2019+"),
                _ => throw new NotSupportedException($"Unknown operator {node.Operation}.")
            };
        }

        public AnyValue Visit(CalcContext context, BinaryNode node)
        {
            var leftArg = node.LeftExpression.Accept(context, this);
            var rightArg = node.RightExpression.Accept(context, this);

            return node.Operation switch
            {
                BinaryOp.Range => AnyValue.ReferenceRange(leftArg, rightArg, context),
                BinaryOp.Union => AnyValue.ReferenceUnion(leftArg, rightArg),
                BinaryOp.Intersection => throw new NotImplementedException("Evaluation of range intersection operator is not implemented."),
                BinaryOp.Concat => AnyValue.Concat(leftArg, rightArg, context),
                BinaryOp.Add => AnyValue.BinaryPlus(leftArg, rightArg, context),
                BinaryOp.Sub => AnyValue.BinaryMinus(leftArg, rightArg, context),
                BinaryOp.Mult => AnyValue.BinaryMult(leftArg, rightArg, context),
                BinaryOp.Div => AnyValue.BinaryDiv(leftArg, rightArg, context),
                BinaryOp.Exp => AnyValue.BinaryExp(leftArg, rightArg, context),
                BinaryOp.Lt => AnyValue.IsLessThan(leftArg, rightArg, context),
                BinaryOp.Lte => AnyValue.IsLessThanOrEqual(leftArg, rightArg, context),
                BinaryOp.Eq => AnyValue.IsEqual(leftArg, rightArg, context),
                BinaryOp.Neq => AnyValue.IsNotEqual(leftArg, rightArg, context),
                BinaryOp.Gte => AnyValue.IsGreaterThanOrEqual(leftArg, rightArg, context),
                BinaryOp.Gt => AnyValue.IsGreaterThan(leftArg, rightArg, context),
                _ => throw new NotSupportedException($"Unknown operator {node.Operation}.")
            };
        }

        public AnyValue Visit(CalcContext context, FunctionNode functionNode)
        {
            if (!_functions.TryGetFunc(functionNode.Name, out var fn))
                return XLError.NameNotRecognized;

            var parameters = functionNode.Parameters;
            var pool = _argsPool.Rent(parameters.Count);
            var args = new Span<AnyValue>(pool, 0, parameters.Count);
            try
            {
                for (var i = 0; i < parameters.Count; ++i)
                    args[i] = parameters[i].Accept(context, this);

                return !context.IsArrayCalculation
                    ? fn.CallFunction(context, args)
                    : fn.CallAsArray(context, args);
            }
            finally
            {
                _argsPool.Return(pool);
            }
        }

        public AnyValue Visit(CalcContext context, ReferenceNode node)
        {
            return node.GetReference(context);
        }

        public AnyValue Visit(CalcContext context, NameNode node)
        {
            return node.GetValue(context.Worksheet, context.CalcEngine);
        }

        public AnyValue Visit(CalcContext context, NotSupportedNode node)
            => throw new NotImplementedException($"Evaluation of {node.FeatureName} is not implemented.");

        public AnyValue Visit(CalcContext context, StructuredReferenceNode node)
        {
            // We don't support external links
            if (node.Prefix is not null)
                return XLError.CellReference;

            if (!TryGetTable(context, node.Table, out var table))
                return XLError.CellReference;

            var area = table.Area;
            if (!TryGetColumn(table, node.FirstColumn, area.LeftColumn, out var colStart))
                return XLError.CellReference;

            if (!TryGetColumn(table, node.LastColumn, area.RightColumn, out var colEnd))
                return XLError.CellReference;

            if (colStart > colEnd)
                (colEnd, colStart) = (colStart, colEnd);

            // Row range is always continuous, so the result is an area. [[#Header],[#Totals]] is
            // not allowed by grammar.
            if (!TryGetRows(context, table, node.Area, out var rowStart, out var rowEnd, out var error))
                return error;

            var range = new XLSheetRange(rowStart, colStart, rowEnd, colEnd);
            return new Reference(XLRangeAddress.FromSheetRange(context.Worksheet, range));

            static bool TryGetTable(CalcContext context, string? tableName, [NotNullWhen(true)] out XLTable? table)
            {
                // table-less references are allowed only in a table area. Excel doesn't allow
                // to set it in GUI, but interprets such situation as #REF!.
                if (tableName is not null)
                {
                    return context.Workbook.TryGetTable(tableName, out table);
                }

                // Avoid LINQ allocation.
                var formulaPoint = context.FormulaSheetPoint;
                foreach (var sheetTable in context.Worksheet.Tables)
                {
                    if (sheetTable.Area.Contains(formulaPoint))
                    {
                        table = sheetTable;
                        return true;
                    }
                }

                table = null;
                return false;
            }

            static bool TryGetColumn(XLTable table, string? column, int defaultColumn, out int columnNo)
            {
                if (column is null)
                {
                    columnNo = defaultColumn;
                    return true;
                }

                if (!table.FieldNames.TryGetValue(column, out var field))
                {
                    columnNo = default;
                    return false;
                }

                columnNo = field.Index + table.Area.LeftColumn;
                return true;
            }

            static bool TryGetRows(CalcContext context, XLTable table, StructuredReferenceArea tableArea,
                out int rowStartNo, out int rowEndNo, out XLError error)
            {
                var area = table.Area;
                var dataEndRowNo = table.ShowTotalsRow ? area.BottomRow - 1 : area.BottomRow;
                switch (tableArea)
                {
                    case StructuredReferenceArea.None:
                    case StructuredReferenceArea.Data:
                        rowStartNo = area.TopRow + 1;
                        rowEndNo = dataEndRowNo;
                        break;
                    case StructuredReferenceArea.Headers:
                        rowStartNo = area.TopRow;
                        rowEndNo = area.TopRow;
                        break;
                    case StructuredReferenceArea.Headers | StructuredReferenceArea.Data:
                        rowStartNo = area.TopRow;
                        rowEndNo = dataEndRowNo;
                        break;
                    case StructuredReferenceArea.Totals:
                        var hasTotals = table.ShowTotalsRow;
                        if (!hasTotals)
                        {
                            rowStartNo = rowEndNo = default;
                            error = XLError.CellReference;
                            return false;
                        }

                        rowStartNo = area.BottomRow;
                        rowEndNo = area.BottomRow;
                        break;
                    case StructuredReferenceArea.Totals | StructuredReferenceArea.Data:
                        rowStartNo = area.TopRow + 1;
                        rowEndNo = area.BottomRow;
                        break;
                    case StructuredReferenceArea.All:
                        rowStartNo = area.TopRow;
                        rowEndNo = area.BottomRow;
                        break;
                    case StructuredReferenceArea.ThisRow:
                        var thisRow = context.FormulaSheetPoint.Row;
                        if (area.TopRow >= thisRow || dataEndRowNo < thisRow)
                        {
                            rowStartNo = rowEndNo = default;
                            error = XLError.IncompatibleValue;
                            return false;
                        }

                        rowStartNo = thisRow;
                        rowEndNo = thisRow;
                        break;
                    default:
                        throw new NotSupportedException($"Unexpected value {tableArea}.");
                }

                error = default;
                return true;
            }
        }

        public AnyValue Visit(CalcContext context, PrefixNode node)
            => throw new InvalidOperationException("Node should never be visited.");

        public AnyValue Visit(CalcContext context, FileNode node)
            => throw new InvalidOperationException("Node should never be visited.");
    }
}
