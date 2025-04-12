using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using ClosedXML.IO.CodeGen.Model.TopLevel;

namespace ClosedXML.IO.CodeGen.Model;

/// <summary>
/// A representation of a one XSD file.
/// </summary>
public class Schema : INode
{
    /// <summary>
    /// Imports in the file.
    /// </summary>
    public List<ImportElement> Imports { get; } = [];

    /// <summary>
    /// One of <c>xsd:attributeGroup</c>, <c>xsd:complexType</c>, <c>xsd:element</c> or <c>xsd:simpleType</c>.
    /// </summary>
    public List<object> Entries { get; } = [];

    T INode.Accept<T>(IXsdVisitor<T> visitor)
    {
        return visitor.Visit(this);
    }

    internal bool TryGetComplexType(string complexTypeName, [NotNullWhen(true)] out ComplexType? complexType)
    {
        complexType = Entries.OfType<ComplexType>().SingleOrDefault(x => x.Name == complexTypeName);
        return complexType is not null;
    }
}
