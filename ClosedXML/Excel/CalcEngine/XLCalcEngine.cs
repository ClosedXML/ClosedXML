using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ClosedXML.Excel.CalcEngine
{
    internal class XLCalcEngine : CalcEngine
    {
        public XLCalcEngine(CultureInfo culture) : base(culture)
        { }

        /// <summary>
        /// Get cells that could be used as input of a formula, that could affect the calculated value.
        /// </summary>
        /// <remarks>Doesn't work for ranges determined by reference functions and reference operators, e.g. <c>A1:IF(SomeCondition,B1,C1)</c>.</remarks>
        /// <param name="expression">Formula to analyze.</param>
        /// <param name="worksheet">Worksheet used for ranges without sheet.</param>
        /// <param name="uniqueCells">All cells (including newly created blank ones) that are referenced in the formula.</param>
        /// <returns>.</returns>
        public bool TryGetPrecedentCells(string expression, XLWorksheet worksheet, out ICollection<XLCell> uniqueCells)
        {
            // This sucks and doesn't work for adding/removing named ranges/worksheets. Also, it creates new cells for all found ranges.
            if (string.IsNullOrWhiteSpace(expression))
            {
                uniqueCells = System.Array.Empty<XLCell>();
                return true;
            }

            var remotelyReliable = TryGetPrecedentAreas(expression, worksheet, out var precedentAreas);
            if (!remotelyReliable)
            {
                uniqueCells = null;
                return false;
            }
            var visitedCells = new HashSet<IXLAddress>(new XLAddressComparer(true));

            var precedentCells = new XLCells(usedCellsOnly: false, XLCellsUsedOptions.Contents);
            foreach (var precedentArea in precedentAreas)
                precedentCells.Add(precedentArea);

            uniqueCells = new List<XLCell>();
            foreach (var cell in precedentCells)
            {
                if (!visitedCells.Contains(cell.Address))
                {
                    visitedCells.Add(cell.Address);
                    uniqueCells.Add(cell);
                }
            }

            return true;
        }

        private bool TryGetPrecedentAreas(string expression, XLWorksheet worksheet, out ICollection<XLRangeAddress> precedentAreas)
        {
            var formula = Parse(expression);
            var ctx = new PrecedentAreasContext(worksheet);
            var rootValue = formula.AstRoot.Accept(ctx, FormulaRangesVisitor.Default);
            if (ctx.HasReferenceErrors/* || ctx.UsesNamedRanges */)
            {
                precedentAreas = null;
                return false;
            }

            if (rootValue.TryPickT0(out var rootReference, out var _))
                ctx.AddReference(rootReference);

            precedentAreas = ctx.FoundReferences
                .SelectMany(x => x.Areas)
                .Select(referenceArea => referenceArea.Worksheet is null
                    ? referenceArea.WithWorksheet(worksheet)
                    : referenceArea)
                .ToList();
            return true;
        }

        private class PrecedentAreasContext
        {
            public PrecedentAreasContext(XLWorksheet worksheet)
            {
                Worksheet = worksheet;
                FoundReferences = new List<Reference>();
            }

            public XLWorksheet Worksheet { get; }

            public List<Reference> FoundReferences { get; }

            /// <summary>
            /// Unable to determine all references, e.g. sheet doesn't exist.
            /// </summary>
            public bool HasReferenceErrors { get; set; }

            public bool UsesNamedRanges { get; set; }

            public void AddReference(Reference reference) => FoundReferences.Add(reference);
        }

        /// <summary>
        /// Get all ranges in the formula. Note that just because range
        /// is in the formula, it doesn't mean it is actually used during evaluation.
        /// Because named ranges can change, the result might change between visits.
        /// </summary>
        private class FormulaRangesVisitor : IFormulaVisitor<PrecedentAreasContext, OneOf<Reference, XLError>>
        {
            public static readonly FormulaRangesVisitor Default = new();

            public OneOf<Reference, XLError> Visit(PrecedentAreasContext ctx, ReferenceNode node)
            {
                if (node.Prefix is null)
                    return new Reference(new XLRangeAddress(null, node.Address));

                if (ctx.Worksheet.Workbook.TryGetWorksheet(node.Prefix?.Sheet, out var ws))
                    return new Reference(new XLRangeAddress((XLWorksheet)ws, node.Address));

                ctx.HasReferenceErrors = true;
                return XLError.CellReference;
            }

            public OneOf<Reference, XLError> Visit(PrecedentAreasContext ctx, NameNode node)
            {
                ctx.UsesNamedRanges = true;

                if (!node.TryGetNameRange(ctx.Worksheet, out var range))
                    return XLError.NameNotRecognized;

                // TODO: This ignores all other ways a name could reference other cells, like A1+5
                if (!range.IsValid)
                {
                    ctx.HasReferenceErrors = true;
                    return XLError.CellReference;
                }

                return new Reference(range.Ranges);

            }

            public OneOf<Reference, XLError> Visit(PrecedentAreasContext ctx, BinaryNode node)
            {
                var leftArg = node.LeftExpression.Accept(ctx, this);

                var rightArg = node.RightExpression.Accept(ctx, this);

                var isLeftReference = leftArg.TryPickT0(out var leftReference, out var leftError);
                var isRightReference = rightArg.TryPickT0(out var rightReference, out var rightError);

                if (!isLeftReference && !isRightReference)
                    return XLError.CellReference;

                if (isLeftReference && !isRightReference)
                {
                    ctx.AddReference(leftReference);
                    return rightError;
                }

                if (!isLeftReference && isRightReference)
                {
                    ctx.AddReference(rightReference);
                    return leftError;
                }

                // Don't add resulting reference into the ctx here, because it still might be turned into an error later (some ranges have many operations A1:B5:C3)
                return node.Operation switch
                {
                    BinaryOp.Range => Reference.RangeOp(leftReference, rightReference, ctx.Worksheet),
                    BinaryOp.Union => Reference.UnionOp(leftReference, rightReference),
                    BinaryOp.Intersection => throw new NotImplementedException("Range intersection not implemented."),
                    _ => XLError.CellReference // Binary operation on reference arguments
                };
            }

            public OneOf<Reference, XLError> Visit(PrecedentAreasContext ctx, ScalarNode node)
            {
                return XLError.CellReference;
            }

            public OneOf<Reference, XLError> Visit(PrecedentAreasContext ctx, UnaryNode node)
            {
                var value = node.Expression.Accept(ctx, this);
                if (!value.TryPickT0(out var reference, out var error))
                    return error;
                ctx.AddReference(reference);
                return XLError.CellReference;
            }

            public OneOf<Reference, XLError> Visit(PrecedentAreasContext ctx, FunctionNode node)
            {
                foreach (var param in node.Parameters)
                {
                    var paramResult = param.Accept(ctx, this);
                    if (paramResult.TryPickT0(out var reference, out _))
                        ctx.AddReference(reference);
                }
                return XLError.CellReference;
            }

            public OneOf<Reference, XLError> Visit(PrecedentAreasContext ctx, NotSupportedNode node)
            {
                return XLError.CellReference;
            }

            public OneOf<Reference, XLError> Visit(PrecedentAreasContext ctx, StructuredReferenceNode node)
            {
                throw new NotImplementedException("Structured references are not implemented.");
            }

            public OneOf<Reference, XLError> Visit(PrecedentAreasContext ctx, PrefixNode node)
            {
                throw new InvalidOperationException("PrefixNode shouldn't be visited.");
            }

            public OneOf<Reference, XLError> Visit(PrecedentAreasContext ctx, FileNode node)
            {
                throw new InvalidOperationException("FileNode shouldn't be visited.");
            }
        }
    }
}
