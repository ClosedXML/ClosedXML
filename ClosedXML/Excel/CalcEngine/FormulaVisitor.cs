namespace ClosedXML.Excel.CalcEngine
{
    internal interface IFormulaVisitor<TContext, TResult>
    {
        public TResult Visit(TContext context, Expression node);

        public TResult Visit(TContext context, UnaryExpression node);

        public TResult Visit(TContext context, BinaryExpression node);

        public TResult Visit(TContext context, FunctionExpression node);

        public TResult Visit(TContext context, VariableExpression node);

        public TResult Visit(TContext context, XObjectExpression node);

        public TResult Visit(TContext context, EmptyValueExpression node);

        public TResult Visit(TContext context, ErrorExpression node);

        public TResult Visit(TContext context, NotSupportedNode node);

        public TResult Visit(TContext context, ReferenceNode node);

        public TResult Visit(TContext context, StructuredReferenceNode node);
    }
}
