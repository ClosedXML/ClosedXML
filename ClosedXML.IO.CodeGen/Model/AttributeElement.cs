using System.Diagnostics;

namespace ClosedXML.IO.CodeGen.Model;

/// <summary>
/// <![CDATA[<xsd:attribute>]]> inside <![CDATA[<xsd:complexType>]]> or <![CDATA[<xsd:attributeGroup>]]>
/// <example>
/// <code><![CDATA[
/// <xsd:attribute name="level" type="xsd:unsignedInt" use="optional" default="0"/>
/// ]]></code>
/// </example>
/// </summary>
public class AttributeElement : INode
{
    /// <summary>
    /// Name is technically optional in ref attribute:
    /// <code>
    ///   <![CDATA[<xsd:attribute ref="r:id" use="optional"/>]]>
    /// </code>
    /// </summary>
    public required string? Name { get; set; }

    public required string? RefName { get; set; }

    public required string? Type { get; set; }

    public AttributeUseType Use { get; set; }

    public string? DefaultValue { get; set; }

    public T Accept<T>(IXsdVisitor<T> visitor)
    {
        return visitor.Visit(this);
    }

    internal void Generate(CodeBuilder code)
    {
        Debug.Assert(Name is not null);
        Debug.Assert(Type is not null);
        code.WriteIndent().Append("var ").AppendVariable(Name).Append(" = ").AppendSimpleTypeMethod(this);
        if (DefaultValue is not null)
            code.Append(" ?? ").Append(DefaultValue);
        code.Append(";").EndLine();
    }
}
