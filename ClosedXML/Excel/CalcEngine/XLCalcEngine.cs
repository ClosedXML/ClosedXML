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
        /// Get a collection of cell ranges included into the expression. Order is not preserved.
        /// </summary>
        /// <param name="worksheet">Worksheet used for ranges without sheet.</param>
        /// <param name="expression">Formula to analyze.</param>
        /// <returns>Collection of ranges included into the expression. Each range address has worksheet.</returns>
        public IEnumerable<IXLRangeAddress> GetPrecedentRanges(XLWorksheet worksheet, string expression)
        {
            var node = Parse(expression);
            var references = new List<Reference>();
            node.Accept(new KeyValuePair<XLWorksheet, List<Reference>>(worksheet, references), FormulaRangesVisitor.Default);

            var visitedRanges = new HashSet<IXLRangeAddress>(new XLRangeAddressComparer(true));
            foreach (var referenceArea in references.SelectMany(x => x.Areas))
            {
                var area = referenceArea.Worksheet is null ? referenceArea.WithWorksheet(worksheet) : referenceArea;
                if (!visitedRanges.Contains(area))
                {
                    visitedRanges.Add(area);
                    yield return area;
                }
            }
        }

        /// <summary>
        /// Get cells that could be used as input of a formula, that could affect the calculated value.
        /// </summary>
        /// <param name="worksheet">Worksheet used for ranges without sheet.</param>
        /// <param name="expression">Formula to analyze.</param>
        /// <returns></returns>
        public IEnumerable<IXLCell> GetPrecedentCells(XLWorksheet worksheet, string expression)
        {
            if (!String.IsNullOrWhiteSpace(expression))
            {
                var node = Parse(expression);
                var ranges = new List<Reference>();
                node.Accept(new KeyValuePair<XLWorksheet, List<Reference>>(worksheet, ranges), FormulaRangesVisitor.Default);

                var wb = worksheet.Workbook;
                var visitedCells = new HashSet<IXLAddress>(new XLAddressComparer(true));

                // TODO: Change semantic of this function so we only return used cells, much more performant
                // I guess I should use some XLCellsUserOptions, but I have no idea which one and conditions are not there anyway.
                var cells = new XLCells(usedCellsOnly: false, XLCellsUsedOptions.Contents);
                foreach (var usedRange in ranges.SelectMany(r => r.Areas))
                    cells.Add(usedRange.Worksheet is null ? usedRange.WithWorksheet(worksheet) : usedRange);

                foreach (var cell in cells)
                {
                    if (!visitedCells.Contains(cell.Address))
                    {
                        visitedCells.Add(cell.Address);
                        yield return cell;
                    }
                }
            }
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

            return new CellRangeReference(range, this);
        }

        /// <summary>
        /// Get all ranges in the formula. Note that just because range
        /// is in the formula, it doesn't mean it is actually used during evaluation.
        /// Because named ranges can change, the result might change between visits.
        /// </summary>
        private class FormulaRangesVisitor : IFormulaVisitor<KeyValuePair<XLWorksheet, List<Reference>>, OneOf<Reference, Error1>>
        {
            public readonly static FormulaRangesVisitor Default = new();

            public OneOf<Reference, Error1> Visit(KeyValuePair<XLWorksheet, List<Reference>> context, ReferenceNode node)
            {
                if (node.Type == ReferenceItemType.NamedRange)
                {
                    // TODO: Cleanup and error checking
                    if (!TryGetNamedRange(node.Address, context.Key, out var namedRange))
                        return Error1.Name;

                    if (!namedRange.IsValid)
                        return Error1.Ref;

                    var rangeAddresses = namedRange.Ranges.Select(r => r.RangeAddress).Cast<XLRangeAddress>().ToList();
                    if (rangeAddresses.Count < 1)
                        throw new NotImplementedException("I guess return error?");
                    return new Reference(rangeAddresses);
                }

                var sheetName = node.Prefix?.Sheet;
                var rangeAddress = sheetName is not null && context.Key.Workbook.TryGetWorksheet(sheetName, out var ws)
                    ? new XLRangeAddress((XLWorksheet)ws, node.Address)
                    : new XLRangeAddress(null, node.Address);

                return new Reference(rangeAddress);
            }

            public OneOf<Reference, Error1> Visit(KeyValuePair<XLWorksheet, List<Reference>> context, BinaryExpression node)
            {
                var leftArg = node.LeftExpression.Accept(context, this);

                var rightArg = node.RightExpression.Accept(context, this);

                var isLeftReference = leftArg.TryPickT0(out var leftReference, out var leftError);
                var isRightReference = rightArg.TryPickT0(out var rightReference, out var rightError);

                if (!isLeftReference && !isRightReference)
                    return Error1.Ref;

                if (isLeftReference && !isRightReference)
                {
                    context.Value.Add(leftReference);
                    return rightError;
                }

                if (!isLeftReference && isRightReference)
                {
                    context.Value.Add(rightReference);
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
                return Error1.Ref;
            }

            public OneOf<Reference, Error1> Visit(KeyValuePair<XLWorksheet, List<Reference>> context, ScalarNode node)
            {
                return Error1.Ref;
            }

            public OneOf<Reference, Error1> Visit(KeyValuePair<XLWorksheet, List<Reference>> context, UnaryExpression node)
            {
                var value = node.Expression.Accept(context, this);
                if (!value.TryPickT0(out var reference, out var error))
                    return error;
                context.Value.Add(reference);
                return Error1.Ref;
            }

            public OneOf<Reference, Error1> Visit(KeyValuePair<XLWorksheet, List<Reference>> context, FunctionExpression node)
            {
                foreach (var param in node.Parameters)
                {
                    var paramResult = param.Accept(context, this);
                    if (paramResult.TryPickT0(out var reference, out var _))
                        context.Value.Add(reference);
                }
                return Error1.Ref;
            }

            public OneOf<Reference, Error1> Visit(KeyValuePair<XLWorksheet, List<Reference>> context, EmptyArgumentNode node)
            {
                return Error1.Ref;
            }

            public OneOf<Reference, Error1> Visit(KeyValuePair<XLWorksheet, List<Reference>> context, ErrorExpression node)
            {
                return Error1.Ref;
            }

            public OneOf<Reference, Error1> Visit(KeyValuePair<XLWorksheet, List<Reference>> context, NotSupportedNode node)
            {
                return Error1.Ref;
            }

            public OneOf<Reference, Error1> Visit(KeyValuePair<XLWorksheet, List<Reference>> context, StructuredReferenceNode node)
            {
                throw new NotImplementedException("Shouldn't be visited.");
            }

            public OneOf<Reference, Error1> Visit(KeyValuePair<XLWorksheet, List<Reference>> context, PrefixNode node)
            {
                throw new InvalidOperationException("Shouldn't be visited.");
            }

            public OneOf<Reference, Error1> Visit(KeyValuePair<XLWorksheet, List<Reference>> context, FileNode node)
            {
                throw new InvalidOperationException("Shouldn't be visited.");
            }
        }
    }
}
