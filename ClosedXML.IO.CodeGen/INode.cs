namespace ClosedXML.IO.CodeGen;

/// <summary>
/// A visitor pattern node, used by a <see cref="IXsdVisitor{T}"/>.
/// </summary>
public interface INode
{
    T Accept<T>(IXsdVisitor<T> visitor);
}
