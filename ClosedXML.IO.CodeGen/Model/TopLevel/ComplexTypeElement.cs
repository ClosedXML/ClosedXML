using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen.Model.TopLevel;

/// <summary>
/// <c><![CDATA[<xsd:complexType/>]]></c> inside <c><![CDATA[<xsd:schema/>]]></c>. It doesn't have
/// any elements, only attributes.
/// </summary>
public class ComplexTypeElement : ComplexType
{
    internal override List<Variable> GenerateParseMethod(CodeBuilder code)
    {
        // Attributes are already parsed by the ComplexType.GenerateParseMethod
        return [];
    }
}
