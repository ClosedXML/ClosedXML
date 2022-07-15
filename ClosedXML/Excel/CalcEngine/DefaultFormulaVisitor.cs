using System.Linq;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// A default visitor that copies a formula.
    /// </summary>
    internal class DefaultFormulaVisitor<TContext> : IFormulaVisitor<TContext, ExpressionBase>
    {
        public virtual ExpressionBase Visit(TContext context, UnaryExpression node)
        {
            var acceptedArgument = (Expression)node.Expression.Accept(context, this);
            return !ReferenceEquals(acceptedArgument, node.Expression)
                ? new UnaryExpression(node.Operation, acceptedArgument)
                : node;
        }

        public virtual ExpressionBase Visit(TContext context, BinaryExpression node)
        {
            var acceptedLeftArgument = (Expression)node.LeftExpression.Accept(context, this);
            var acceptedRightArgument = (Expression)node.RightExpression.Accept(context, this);
            return !ReferenceEquals(acceptedLeftArgument, node.LeftExpression) || !ReferenceEquals(acceptedRightArgument, node.RightExpression)
                ? new BinaryExpression(node.Operation, acceptedLeftArgument, acceptedRightArgument)
                : node;
        }

        public virtual ExpressionBase Visit(TContext context, FunctionExpression node)
        {
            var acceptedParameters = node.Parameters.Select(p => p.Accept(context, this)).Cast<Expression>().ToList();
            return node.Parameters.Zip(acceptedParameters, (param, acceptedParam) => !ReferenceEquals(param, acceptedParam)).Any()
                ? new FunctionExpression(node.Prefix, node.FunctionDefinition, acceptedParameters)
                : node;
        }

        public virtual ExpressionBase Visit(TContext context, XObjectExpression node) => node;

        public virtual ExpressionBase Visit(TContext context, EmptyValueExpression node) => node;

        public virtual ExpressionBase Visit(TContext context, ScalarNode node) => node;

        public virtual ExpressionBase Visit(TContext context, ErrorExpression node) => node;

        public virtual ExpressionBase Visit(TContext context, NotSupportedNode node) => node;

        public virtual ExpressionBase Visit(TContext context, ReferenceNode referenceNode)
        {
            var acceptedPrefix = referenceNode.Prefix?.Accept(context, this);
            return !ReferenceEquals(acceptedPrefix, referenceNode.Prefix)
                ? new ReferenceNode((PrefixNode)acceptedPrefix, referenceNode.Type, referenceNode.Address)
                : referenceNode;
        }

        public virtual ExpressionBase Visit(TContext context, StructuredReferenceNode node) => node;

        public virtual ExpressionBase Visit(TContext context, PrefixNode prefix)
        {
            var acceptedFile = prefix.File?.Accept(context, this);
            return !ReferenceEquals(acceptedFile, prefix.File)
                ? new PrefixNode((FileNode)acceptedFile, prefix.Sheet, prefix.FirstSheet, prefix.LastSheet)
                : prefix;
        }

        public virtual ExpressionBase Visit(TContext context, FileNode node) => node;
    }
}
