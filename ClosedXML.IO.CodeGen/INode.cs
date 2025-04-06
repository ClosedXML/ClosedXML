namespace ClosedXML.IO.CodeGen;

/// <summary>
/// A node visited by a <see cref="IXsdVisitor{TResult}"/>.
/// </summary>
public interface INode
{
    T Accept<T>(IXsdVisitor<T> visitor);
}
