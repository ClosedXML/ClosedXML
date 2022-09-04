using System.Linq;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// A default visitor that copies a formula.
    /// </summary>
    internal class DefaultFormulaVisitor<TContext> : IFormulaVisitor<TContext, AstNode>
    {
        public virtual AstNode Visit(TContext context, UnaryNode node)
        {
            var acceptedArgument = (ValueNode)node.Expression.Accept(context, this);
            return !ReferenceEquals(acceptedArgument, node.Expression)
                ? new UnaryNode(node.Operation, acceptedArgument)
                : node;
        }

        public virtual AstNode Visit(TContext context, BinaryNode node)
        {
            var acceptedLeftArgument = (ValueNode)node.LeftExpression.Accept(context, this);
            var acceptedRightArgument = (ValueNode)node.RightExpression.Accept(context, this);
            return !ReferenceEquals(acceptedLeftArgument, node.LeftExpression) || !ReferenceEquals(acceptedRightArgument, node.RightExpression)
                ? new BinaryNode(node.Operation, acceptedLeftArgument, acceptedRightArgument)
                : node;
        }

        public virtual AstNode Visit(TContext context, FunctionNode node)
        {
            var acceptedParameters = node.Parameters.Select(p => p.Accept(context, this)).Cast<ValueNode>().ToList();
            return node.Parameters.Zip(acceptedParameters, (param, acceptedParam) => !ReferenceEquals(param, acceptedParam)).Any()
                ? new FunctionNode(node.Prefix, node.Name, acceptedParameters)
                : node;
        }

        public virtual AstNode Visit(TContext context, ScalarNode node) => node;

        public virtual AstNode Visit(TContext context, NotSupportedNode node) => node;

        public virtual AstNode Visit(TContext context, ReferenceNode referenceNode)
        {
            var acceptedPrefix = referenceNode.Prefix?.Accept(context, this);
            return !ReferenceEquals(acceptedPrefix, referenceNode.Prefix)
                ? new ReferenceNode((PrefixNode)acceptedPrefix, referenceNode.Type, referenceNode.Address)
                : referenceNode;
        }

        public virtual AstNode Visit(TContext context, NameNode nameNode)
        {
            var acceptedPrefix = nameNode.Prefix?.Accept(context, this);
            return !ReferenceEquals(acceptedPrefix, nameNode.Prefix)
                ? new NameNode((PrefixNode)acceptedPrefix, nameNode.Name)
                : nameNode;
        }

        public virtual AstNode Visit(TContext context, StructuredReferenceNode node) => node;

        public virtual AstNode Visit(TContext context, PrefixNode prefix)
        {
            var acceptedFile = prefix.File?.Accept(context, this);
            return !ReferenceEquals(acceptedFile, prefix.File)
                ? new PrefixNode((FileNode)acceptedFile, prefix.Sheet, prefix.FirstSheet, prefix.LastSheet)
                : prefix;
        }

        public virtual AstNode Visit(TContext context, FileNode node) => node;
    }
}
