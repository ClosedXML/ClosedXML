using System;
using System.Collections.Generic;
using System.Linq;
using OneOf;

namespace ClosedXML.Excel.CalcEngine
{
    internal class XLCalcEngine : CalcEngine
    {
        private readonly IXLWorksheet _ws;
        private readonly XLWorkbook _wb;

        public XLCalcEngine()
        { }

        public XLCalcEngine(XLWorkbook wb)
        {
            _wb = wb;
        }

        public XLCalcEngine(IXLWorksheet ws) : this(ws.Workbook)
        {
            _ws = ws;
        }

        /// <summary>
        /// Get a best guess of collection of cell ranges in the formula. Order is not preserved.
        /// </summary>
        /// <remarks>Doesn't work for ranges determined by reference functions and reference operators, e.g. <c>A1:IF(SomeCondition,B1,C1)</c>.</remarks>
        /// <param name="worksheet">Worksheet used for ranges without sheet.</param>
        /// <param name="expression">Formula to analyze.</param>
        /// <returns>Collection of range addresses that are referenced in the formula. All addresses have specified worksheet.</returns>
        public IEnumerable<IXLRangeAddress> GetPrecedentRanges(string expression, XLWorksheet worksheet)
        {
            // TODO: Unused function... delete?
            var remotelyReliable = TryGetPrecedentAreas(expression, worksheet, out var referencedAreas);
            var visitedRanges = new HashSet<IXLRangeAddress>(new XLRangeAddressComparer(true));
            foreach (var referenceArea in referencedAreas)
            {
                if (!visitedRanges.Contains(referenceArea))
                {
                    visitedRanges.Add(referenceArea);
                    yield return referenceArea;
                }
            }
        }

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

        public override object GetExternalObject(string identifier)
        {
            if (identifier.Contains("!") && _wb != null)
            {
                var referencedSheetNames = identifier.Split(':')
                    .Select(part =>
                    {
                        if (part.Contains("!"))
                            return part.Substring(0, part.LastIndexOf('!')).ToLower();
                        else
                            return null;
                    })
                    .Where(sheet => sheet != null)
                    .Distinct()
                    .ToList();

                if (referencedSheetNames.Count == 0)
                    return GetCellRangeReference(_ws.Range(identifier));
                else if (referencedSheetNames.Count > 1)
                    throw new ArgumentOutOfRangeException(referencedSheetNames.Last(), "Cross worksheet references may references no more than 1 other worksheet");
                else
                {
                    if (!_wb.TryGetWorksheet(referencedSheetNames.Single(), out IXLWorksheet worksheet))
                        throw new ArgumentOutOfRangeException(referencedSheetNames.Single(), "The required worksheet cannot be found");

                    identifier = identifier.ToLower().Replace(string.Format("{0}!", worksheet.Name.ToLower()), "");

                    return GetCellRangeReference(worksheet.Range(identifier));
                }
            }
            else if (_ws != null)
            {
                if (TryGetNamedRange(identifier, _ws, out IXLNamedRange namedRange))
                {
                    var references = (namedRange as XLNamedRange).RangeList.Select(r =>
                        XLHelper.IsValidRangeAddress(r)
                            ? GetCellRangeReference(_ws.Workbook.Range(r))
                            : new XLCalcEngine(_ws).Evaluate(r.ToString())
                        );
                    if (references.Count() == 1)
                        return references.Single();
                    return references;
                }

                var range = _ws.Range(identifier);
                if (range is null)
                    throw new ArgumentOutOfRangeException("Not a range nor named range.");

                return GetCellRangeReference(range);
            }
            else if (XLHelper.IsValidRangeAddress(identifier))
                return identifier;
            else
                return null;
        }

        private static bool TryGetNamedRange(string identifier, IXLWorksheet worksheet, out IXLNamedRange namedRange)
        {
            return worksheet.NamedRanges.TryGetValue(identifier, out namedRange) ||
                   worksheet.Workbook.NamedRanges.TryGetValue(identifier, out namedRange);
        }

        private CellRangeReference GetCellRangeReference(IXLRange range)
        {
            if (range == null)
                return null;

            return new CellRangeReference(range);
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

            public bool HasReferenceErrors { get; set; }

            public bool UsesNamedRanges { get; set; }

            public void AddReference(Reference reference) => FoundReferences.Add(reference);
        }

        /// <summary>
        /// Get all ranges in the formula. Note that just because range
        /// is in the formula, it doesn't mean it is actually used during evaluation.
        /// Because named ranges can change, the result might change between visits.
        /// </summary>
        private class FormulaRangesVisitor : IFormulaVisitor<PrecedentAreasContext, OneOf<Reference, Error>>
        {
            public readonly static FormulaRangesVisitor Default = new();

            public OneOf<Reference, Error> Visit(PrecedentAreasContext ctx, ReferenceNode node)
            {
                if (node.Type == ReferenceItemType.NamedRange)
                {
                    ctx.UsesNamedRanges = true;

                    // TODO: Cleanup and error checking
                    if (!TryGetNamedRange(node.Address, ctx.Worksheet, out var namedRange))
                    {
                        return Error.NameNotRecognized;
                    }

                    if (!namedRange.IsValid)
                    {
                        ctx.HasReferenceErrors = true;
                        return Error.CellReference;
                    }

                    var rangeAddresses = namedRange.Ranges.Select(r => r.RangeAddress).Cast<XLRangeAddress>().ToList();
                    if (rangeAddresses.Count < 1)
                        throw new NotImplementedException("I guess return error?");
                    return new Reference(rangeAddresses);
                }

                var sheetName = node.Prefix?.Sheet;
                if (sheetName is not null)
                {
                    if (!ctx.Worksheet.Workbook.TryGetWorksheet(sheetName, out var ws))
                    {
                        ctx.HasReferenceErrors = true;
                        return Error.CellReference;
                    }

                    return new Reference(new XLRangeAddress((XLWorksheet)ws, node.Address));
                }

                return new Reference(new XLRangeAddress(null, node.Address));
            }

            public OneOf<Reference, Error> Visit(PrecedentAreasContext ctx, BinaryExpression node)
            {
                var leftArg = node.LeftExpression.Accept(ctx, this);

                var rightArg = node.RightExpression.Accept(ctx, this);

                var isLeftReference = leftArg.TryPickT0(out var leftReference, out var leftError);
                var isRightReference = rightArg.TryPickT0(out var rightReference, out var rightError);

                if (!isLeftReference && !isRightReference)
                    return Error.CellReference;

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

                // Only result store the place where reference would change to error. Some ranges have many operations A1:B5:C3
                switch (node.Operation)
                {
                    case BinaryOp.Range: return Reference.RangeOp(leftReference, rightReference);
                    case BinaryOp.Union: return Reference.UnionOp(leftReference, rightReference);
                    case BinaryOp.Intersection: throw new NotImplementedException();
                };

                // Binary operation on reference arguments
                return Error.CellReference;
            }

            public OneOf<Reference, Error> Visit(PrecedentAreasContext ctx, ScalarNode node)
            {
                return Error.CellReference;
            }

            public OneOf<Reference, Error> Visit(PrecedentAreasContext ctx, UnaryExpression node)
            {
                var value = node.Expression.Accept(ctx, this);
                if (!value.TryPickT0(out var reference, out var error))
                    return error;
                ctx.AddReference(reference);
                return Error.CellReference;
            }

            public OneOf<Reference, Error> Visit(PrecedentAreasContext ctx, FunctionExpression node)
            {
                foreach (var param in node.Parameters)
                {
                    var paramResult = param.Accept(ctx, this);
                    if (paramResult.TryPickT0(out var reference, out var _))
                        ctx.AddReference(reference);
                }
                return Error.CellReference;
            }

            public OneOf<Reference, Error> Visit(PrecedentAreasContext ctx, EmptyArgumentNode node)
            {
                return Error.CellReference;
            }

            public OneOf<Reference, Error> Visit(PrecedentAreasContext ctx, ErrorExpression node)
            {
                return Error.CellReference;
            }

            public OneOf<Reference, Error> Visit(PrecedentAreasContext ctx, NotSupportedNode node)
            {
                return Error.CellReference;
            }

            public OneOf<Reference, Error> Visit(PrecedentAreasContext ctx, StructuredReferenceNode node)
            {
                throw new NotImplementedException("Shouldn't be visited.");
            }

            public OneOf<Reference, Error> Visit(PrecedentAreasContext ctx, PrefixNode node)
            {
                throw new InvalidOperationException("Shouldn't be visited.");
            }

            public OneOf<Reference, Error> Visit(PrecedentAreasContext ctx, FileNode node)
            {
                throw new InvalidOperationException("Shouldn't be visited.");
            }
        }
    }
}
