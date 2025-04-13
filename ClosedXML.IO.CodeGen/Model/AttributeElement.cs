using System.Diagnostics;
using System;
using System.Collections.Generic;

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
    /// C# keywords. The variables with that name must be escaped, e.g. <c>in</c> must be <c>@in</c>.
    /// </summary>
    private static readonly HashSet<string> Keywords = ["in", "out", "ref"];

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
        var isOptional = Use != AttributeUseType.Required;
        var methodTemplate = code.GetSimpleTypeTemplate(Type, isOptional);
        var readAttrExpression = "var " + EscapeVar(Name) + " = " + string.Format(methodTemplate, Name);
        var readAttrCode = DefaultValue is null
            ? readAttrExpression + ";"
            : readAttrExpression + " ?? " + DefaultValue + ";";
        code.AddLine(readAttrCode);
    }

    private static string EscapeVar(string name)
    {
        return Keywords.Contains(name) ? '@' + name : name;
    }
}
