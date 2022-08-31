namespace ClosedXML.Excel.CalcEngine
{
    internal interface IFormulaVisitor<TContext, TResult>
    {
        public TResult Visit(TContext context, ScalarNode node);

        public TResult Visit(TContext context, UnaryNode node);

        public TResult Visit(TContext context, BinaryNode node);

        public TResult Visit(TContext context, FunctionNode node);

        public TResult Visit(TContext context, XObjectExpression node);

        public TResult Visit(TContext context, EmptyArgumentNode node);

        public TResult Visit(TContext context, ErrorNode node);

        public TResult Visit(TContext context, NotSupportedNode node);

        public TResult Visit(TContext context, ReferenceNode node);

        public TResult Visit(TContext context, StructuredReferenceNode node);

        public TResult Visit(TContext context, PrefixNode node);

        public TResult Visit(TContext context, FileNode node);
    }
}
