using System.Linq;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// A default visitor that copies a formula.
    /// </summary>
    internal class DefaultFormulaVisitor<TContext> : IFormulaVisitor<TContext, Expression>
    {
        public virtual Expression Visit(TContext context, Expression node)
        {
            return node;
        }

        public virtual Expression Visit(TContext context, UnaryExpression node)
        {
            var acceptedArgument = node.Expression.Accept(context, this);
            return !ReferenceEquals(acceptedArgument, node.Expression)
                ? new UnaryExpression(node.Operation, acceptedArgument)
                : node;
        }

        public virtual Expression Visit(TContext context, BinaryExpression node)
        {
            var acceptedLeftArgument = node.LeftExpression.Accept(context, this);
            var acceptedRightArgument = node.RightExpression.Accept(context, this);
            return !ReferenceEquals(acceptedLeftArgument, node.LeftExpression) || !ReferenceEquals(acceptedRightArgument, node.RightExpression)
                ? new BinaryExpression(node.Operation, acceptedLeftArgument, acceptedRightArgument)
                : node;
        }

        public virtual Expression Visit(TContext context, FunctionExpression node)
        {
            var acceptedParameters = node.Parameters.Select(p => p.Accept(context, this)).ToList();
            return node.Parameters.Zip(acceptedParameters, (param, acceptedParam) => !ReferenceEquals(param, acceptedParam)).Any()
                ? new FunctionExpression(node.Prefix, node.FunctionDefinition, acceptedParameters)
                : node;
        }

        public virtual Expression Visit(TContext context, XObjectExpression node)
        {
            return node;
        }

        public virtual Expression Visit(TContext context, EmptyValueExpression node)
        {
            return node;
        }

        public virtual Expression Visit(TContext context, ErrorExpression node)
        {
            return node;
        }

        public virtual Expression Visit(TContext context, NotSupportedNode node)
        {
            return node;
        }

        public virtual Expression Visit(TContext context, ReferenceNode node)
        {
            return node;
        }

        public virtual Expression Visit(TContext context, StructuredReferenceNode node)
        {
            return node;
        }
    }
}
