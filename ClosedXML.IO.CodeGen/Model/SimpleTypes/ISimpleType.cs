namespace ClosedXML.IO.CodeGen.Model.SimpleTypes;

/// <summary>
/// A marker interface for types inside <c><![CDATA[<xsd:simpleType>]]></c>.
/// </summary>
public interface ISimpleType : INode
{
    string Name { get; }
}
