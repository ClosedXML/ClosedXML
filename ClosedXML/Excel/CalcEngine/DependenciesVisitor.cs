using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Extensions;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// <para>
    /// Visit each node and determine all ranges that might affect the formula.
    /// It uses concrete values (e.g. actual range for structured references) and
    /// should be refreshed when structured reference or name is changed in a workbook.
    /// </para>
    /// <para>
    /// The areas found by the visitor shouldn't change when data on a worksheet changes,
    /// so the output is a superset of areas, if necessary.
    /// </para>
    /// <para>
    /// Precedents visitor is not completely accurate, in case of uncertainty, it uses
    /// a larger area. At worst the end result is unnecessary recalculation. For simple
    /// cases, it works fine and freaks like <c>A1:IF(Other!B5,B7,Different!G3)</c>
    /// will be marked as dirty more often than strictly necessary.
    /// </para>
    /// <para>
    /// Each node visitor evaluates, if the output is a reference or a value/array. If
    /// the result is an array, it propagates to upper nodes, where can be things like
    /// range operator.
    /// </para>
    /// </summary>
    internal class DependenciesVisitor : IFormulaVisitor<DependenciesContext, List<XLBookArea>?>
    {
        public List<XLBookArea>? Visit(DependenciesContext context, ScalarNode node)
        {
            // Scalar node can't contain sub-nodes or references.
            return null;
        }

        public List<XLBookArea>? Visit(DependenciesContext context, ArrayNode node)
        {
            // Array node can't contain sub-nodes or references.
            return null;
        }

        public List<XLBookArea>? Visit(DependenciesContext context, UnaryNode node)
        {
            var sheetAreas = node.Expression.Accept(context, this);

            // If the operand of unary node is not a reference -> end immediately,
            // operator can't modify non-reference into a reference.
            if (sheetAreas is null)
                return null;

            // Operand is a reference
            if (node.Operation is UnaryOp.ImplicitIntersection or UnaryOp.SpillRange)
            {
                // Both operators are ignored for now, because *spill* operators
                // is part of dynamic arrays (=ignore until they are implemented)
                // and *Implicit intersection* only makes area smaller and is pretty
                // rare operators = also skip for now (won't affect correctness).

                // The reference must be propagated upward, because there could be
                // a range operator (e.g. `B7:@A1:A5`)
                return sheetAreas;
            }

            // Some other operator is applied to the reference -> reference is converted
            // to an array
            context.AddAreas(sheetAreas);
            return null;
        }

        public List<XLBookArea>? Visit(DependenciesContext context, BinaryNode node)
        {
            // Only range operations transform ranges
            var leftAreas = node.LeftExpression.Accept(context, this);
            var rightAreas = node.RightExpression.Accept(context, this);

            // Reference operation only makes sense, if both sides are references.
            // Otherwise reference operation results into an error.
            if (leftAreas is not null && rightAreas is not null)
            {
                // Both sides are references - calculate new ranges and propagate
                if (node.Operation == BinaryOp.Union)
                {
                    leftAreas.AddRange(rightAreas);
                    return leftAreas;
                }

                if (node.Operation == BinaryOp.Range)
                {
                    var rangeResult = new List<XLBookArea>();

                    // Create a new range from both operands. It must deal with
                    // situation where there are multiple sheets for both operands,
                    // e.g. `IF(G4,Sheet1!A1,Sheet2!A2):IF(H3,Sheet2!C4,Sheet1!C5)`
                    // that creates a valid range.
                    var sheetGroups = leftAreas.Concat(rightAreas)
                        .GroupBy(area => area.Name, XLHelper.SheetComparer);

                    // There is no simple way to go through all paths, so try to find
                    // largest possible ranges that could be the result. For normal
                    // operands(A1:B2:C3), it will work fine and for freaks, it will
                    // find largest possible range that is a superset of actual result.
                    foreach (var sheetGroup in sheetGroups)
                    {
                        var sheetAreas = sheetGroup.ToList();
                        if (sheetAreas.Count == 1)
                            continue;

                        var rangeArea = sheetAreas[0].Area;
                        for (var i = 1; i < sheetAreas.Count; ++i)
                            rangeArea = rangeArea.Range(sheetAreas[i].Area);

                        rangeResult.Add(new XLBookArea(sheetGroup.Key, rangeArea));
                    }

                    // It's enough to return result of range operation. Operands can
                    // be discarded, because they are included in the result.
                    return rangeResult;
                }

                if (node.Operation == BinaryOp.Intersection)
                {
                    // Intersection makes range smaller, so it's rather hard to optimize
                    // areas. We make a special case for the most frequent case.
                    if (leftAreas.Count == 1 && rightAreas.Count == 1)
                    {
                        var leftArea = leftAreas[0];
                        var rightArea = rightAreas[0];
                        var intersection = leftArea.Intersect(rightArea);

                        // Propagate only the intersection, not operands. Even if operands
                        // change, it doesn't affect the formula, because cells outside
                        // intersection are never used.
                        if (intersection is not null)
                            return new List<XLBookArea> { intersection.Value };

                        return null;
                    }

                    // Anything else is too complicated and thus just propagate all references.
                    leftAreas.AddRange(rightAreas);
                    return leftAreas;
                }

                // Operand is not a reference one, so reference is turned to array of values.
                context.AddAreas(leftAreas);
                context.AddAreas(rightAreas);
                return null;
            }

            // Both children aren't references or only one is -> binary operation transforms it
            // to a non-reference, either value or #REF!
            if (leftAreas is not null)
                context.AddAreas(leftAreas);

            if (rightAreas is not null)
                context.AddAreas(rightAreas);

            return null;
        }

        public List<XLBookArea>? Visit(DependenciesContext context, FunctionNode node)
        {
            // According to grammar, ref functions are: CHOOSE, IF, INDEX, INDIRECT, OFFSET
            // Only these functions are allowed to return references, per grammar.
            // However, OFFSET and INDIRECT are volatile function that always have to be
            // recalculated (=are always marked dirty).

            if (XLHelper.FunctionComparer.Equals(node.Name, "IF"))
            {
                // Tested value is not propagated, it's evaluated as an argument
                var testReference = node.Parameters[0].Accept(context, this);
                if (testReference is not null)
                    context.AddAreas(testReference);

                // If argument is reference and test is evaluated to TRUE,
                // the reference is returned => propagate.
                var valueIfTrueReference = node.Parameters[1].Accept(context, this);
                var valueIfFalseReference = node.Parameters.Count == 3
                    ? node.Parameters[2].Accept(context, this)
                    : null;

                if (valueIfFalseReference is not null && valueIfTrueReference is not null)
                {
                    valueIfTrueReference.AddRange(valueIfFalseReference);
                    return valueIfTrueReference;
                }

                return valueIfFalseReference ?? valueIfTrueReference;
            }

            if (XLHelper.FunctionComparer.Equals(node.Name, "INDEX"))
            {
                // Add argument references, INDEX can have 2 or 3 arguments
                for (var i = 1; i < node.Parameters.Count; ++i)
                {
                    var argReference = node.Parameters[i].Accept(context, this);
                    if (argReference is not null)
                        context.AddAreas(argReference);
                }

                // If INDEX function indexes into an area, it returns reference,
                // not a value. Either way, return whole reference that is indexed,
                // even though it's larger than actual function result.
                var arrayReference = node.Parameters[0].Accept(context, this);
                return arrayReference;
            }

            if (XLHelper.FunctionComparer.Equals(node.Name, "CHOOSE"))
            {
                // Index argument is used to select value, so don't propagate
                var indexReference = node.Parameters[0].Accept(context, this);
                if (indexReference is not null)
                    context.AddAreas(indexReference);

                // Any of arguments can be propagated -> propagate all.
                // Initialize list as null to reduce allocations
                List<XLBookArea>? parametersReference = null;
                for (var i = 1; i < node.Parameters.Count; ++i)
                {
                    var parameterReference = node.Parameters[i].Accept(context, this);
                    if (parameterReference is null)
                        continue;

                    if (parametersReference is not null)
                        parametersReference.AddRange(parameterReference);
                    else
                        parametersReference = parameterReference;
                }

                return parametersReference;
            }

            // All other functions can have references as arguments, but not as an output value.
            foreach (var parameterNode in node.Parameters)
            {
                var paramReference = parameterNode.Accept(context, this);
                if (paramReference is not null)
                    context.AddAreas(paramReference);
            }

            return null;
        }

        public List<XLBookArea>? Visit(DependenciesContext context, NotSupportedNode node)
        {
            return null;
        }

        public List<XLBookArea>? Visit(DependenciesContext context, ReferenceNode node)
        {
            var prefix = node.Prefix;
            string sheetName;
            if (prefix is not null)
            {
                // We don't support external references, so there is no way to depend on something
                // in different workbook at the moment.
                if (prefix.File is not null)
                    return null;

                // 3D references are not supported yet, so don't propagate anything.
                if (prefix.FirstSheet is not null || prefix.LastSheet is not null)
                    return null;

                sheetName = prefix.Sheet ?? throw new InvalidOperationException("Prefix doesn't contain sheet.");
            }
            else
            {
                sheetName = context.FormulaArea.Name;
            }

            var anchor = context.FormulaArea.Area.FirstPoint;
            var sheetRange = node.ReferenceArea.ToSheetRange(anchor);
            return new List<XLBookArea> { new(sheetName, sheetRange) };
        }

        public List<XLBookArea>? Visit(DependenciesContext context, NameNode node)
        {
            // External references are not supported for names
            if (node.Prefix?.File is not null)
                return null;

            var name = node.Prefix?.Sheet is { } sheetName
                ? new XLName(sheetName, node.Name)
                : new XLName(node.Name);
            context.AddName(name);

            // First, try to interpret name as a sheet scoped name.
            sheetName = node.Prefix?.Sheet ?? context.FormulaArea.Name;
            if (context.Workbook.TryGetWorksheet(sheetName, out XLWorksheet sheet) &&
                sheet.DefinedNames.TryGetScopedValue(node.Name, out var sheetDefinedName))
            {
                return VisitName(sheetDefinedName);
            }

            // Name is not a sheet scoped one, try workbook scoped one
            if (context.Workbook.DefinedNamesInternal.TryGetScopedValue(node.Name, out var bookNamedRange))
            {
                return VisitName(bookNamedRange!);
            }

            // Name is not found in the workbook
            return null;

            List<XLBookArea>? VisitName(XLDefinedName definedName)
            {
                // The named range is stored as A1 and thus parsed as A1, but should be interpreted as R1C1
                var namedFormula = definedName.RefersTo;
                var ast = context.Workbook.CalcEngine.Parse(namedFormula);
                var nameReferences = ast.AstRoot.Accept(context, this);

                // If the formula returned a reference, propagate it, rather
                // than add to the context (required for `A1:name` ).
                return nameReferences;
            }
        }

        public List<XLBookArea>? Visit(DependenciesContext context, StructuredReferenceNode node)
        {
            // TODO: Structured reference should be evaluated into a reference and propagated.
            return null;
        }

        public List<XLBookArea> Visit(DependenciesContext context, PrefixNode node)
        {
            throw new InvalidOperationException("Should never be called.");
        }

        public List<XLBookArea> Visit(DependenciesContext context, FileNode node)
        {
            throw new InvalidOperationException("Should never be called.");
        }
    }
}
