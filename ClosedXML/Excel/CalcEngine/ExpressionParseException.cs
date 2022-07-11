using Irony.Parsing;
using System;
using System.Text;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// The exception that is thrown when the strings to be parsed to an expression is invalid.
    /// </summary>
    public class ExpressionParseException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the ExpressionParseException class with a 
        /// specified error message.
        /// </summary>
        /// <param name="message">The message that describes the error.</param>
        public ExpressionParseException(string message)
            : base(message)
        {
        }

        internal ExpressionParseException(ParseTreeNode parseNode)
            : base($"Parser error at node {parseNode.Term.Name}, location: {parseNode.Span.Location.Line}:{parseNode.Span.Location.Position}-{parseNode.Span.EndPosition}")
        {
        }
    }
}
