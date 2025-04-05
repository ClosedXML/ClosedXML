using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen.Model.Elements;

/// <summary>
/// A node in a complex type element tree.
/// </summary>
public interface IElementGroup: INode
{
    /// <summary>
    /// Children elements.
    /// </summary>
    public List<IElementGroup> Children { get; }
}
