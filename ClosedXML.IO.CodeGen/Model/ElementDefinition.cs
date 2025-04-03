namespace ClosedXML.IO.CodeGen.Model;

/// <summary>
/// A definition of the element in the root. The element might be referenced in complex types.
/// <code><![CDATA[<xsd:element name="calcChain" type="CT_CalcChain"/>]]></code>
/// </summary>
public class ElementDefinition
{
    /// <summary>
    /// Name of the element.
    /// </summary>
    public required string Name { get; init; }

    /// <summary>
    /// The type name of the element.
    /// </summary>
    public required string TypeName { get; init; }
}
