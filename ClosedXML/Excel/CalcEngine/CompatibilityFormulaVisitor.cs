namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// Old parser was using <see cref="XObjectExpression"/> for references. AST represents references through
    /// other nodes. The visitor replaces a subset of <see cref="ReferenceNode"/> patterns with <see cref="XObjectExpression"/>
    /// so everything works as it did with the old parser. Will be removed once reference operations are implemented.
    /// </summary>
    internal class CompatibilityFormulaVisitor : DefaultFormulaVisitor<object>
    {
        private readonly CalcEngine _calcEngine;

        public CompatibilityFormulaVisitor(CalcEngine calcEngine)
        {
            _calcEngine = calcEngine;
        }

        public override Expression Visit(object context, ReferenceNode referenceNode)
        {
            // Pattern: 'A1' or 'SomeNamedRange' or 'A:Z' or '1:14'
            if (referenceNode.Prefix is null)
            {
                return new XObjectExpression(_calcEngine.GetExternalObject(referenceNode.Address));
            }

            // Pattern: 'Sheet!A1' or 'Sheet!SomeNamedRange' or 'Sheet!A:Z' or 'Sheet!1:14'
            if (referenceNode.Prefix.File is null && referenceNode.Prefix.Sheet is not null)
            {
                return new XObjectExpression(_calcEngine.GetExternalObject(referenceNode.Prefix.Sheet.EscapeSheetName() + "!" + referenceNode.Address));
            }

            return base.Visit(context, referenceNode);
        }

        public override Expression Visit(object context, BinaryExpression binaryNode)
        {
            if (binaryNode.Operation == BinaryOp.Range
                && binaryNode.LeftExpression is ReferenceNode leftReference
                && binaryNode.RightExpression is ReferenceNode rightReference)
            {
                // Pattern A1:B2
                if (leftReference.Prefix is null
                    && leftReference.Type == ReferenceItemType.Cell
                    && rightReference.Prefix is null
                    && rightReference.Type == ReferenceItemType.Cell)
                {
                    return new XObjectExpression(_calcEngine.GetExternalObject(leftReference.Address + ":" + rightReference.Address));
                }

                // Pattern Sheet!A1:B2
                if (leftReference.Prefix is not null
                    && leftReference.Prefix.File is null
                    && leftReference.Prefix.Sheet is not null
                    && leftReference.Type == ReferenceItemType.Cell
                    && rightReference.Prefix is null
                    && rightReference.Type == ReferenceItemType.Cell)
                {
                    return new XObjectExpression(_calcEngine.GetExternalObject(leftReference.Prefix.Sheet.EscapeSheetName() + "!" + leftReference.Address + ":" + rightReference.Address));
                }

                // Pattern SomeSheet!A1:OtherSheet!B2, if from one sheet = OK, otherwise ArgumentOutOfRangeException
                if (leftReference.Prefix is not null
                    && leftReference.Prefix.File is null
                    && leftReference.Prefix.Sheet is not null
                    && leftReference.Type == ReferenceItemType.Cell
                    && rightReference.Prefix is not null
                    && rightReference.Prefix.File is null
                    && rightReference.Type == ReferenceItemType.Cell
                    && rightReference.Prefix.Sheet is not null)
                {
                    return new XObjectExpression(_calcEngine.GetExternalObject(leftReference.Prefix.Sheet.EscapeSheetName() + "!" + leftReference.Address + ":" + rightReference.Prefix.Sheet.EscapeSheetName() + "!" + rightReference.Address));
                }
            }

            return base.Visit(context, binaryNode);
        }
    }
}
