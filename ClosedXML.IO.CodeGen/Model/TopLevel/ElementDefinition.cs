using ClosedXML.IO.CodeGen.Model.Elements;

namespace ClosedXML.IO.CodeGen.Model.TopLevel;

/// <summary>
/// <c><![CDATA[<xsd:element/>]]></c> inside <c><![CDATA[<xsd:schema/>]]></c>.
/// <example>
/// <code><![CDATA[
///   <xsd:element name="calcChain" type="CT_CalcChain"/>
/// ]]></code>
/// </example>
/// </summary>
public class ElementDefinition : IReferencable
{
    /// <summary>
    /// Name of the element. Referenced by <see cref="ElementReference.RefName"/>.
    /// </summary>
    public required string Name { get; init; }

    /// <summary>
    /// The type name of the element.
    /// </summary>
    public required string TypeName { get; init; }
}
