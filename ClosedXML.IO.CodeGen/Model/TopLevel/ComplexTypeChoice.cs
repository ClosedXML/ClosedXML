using ClosedXML.IO.CodeGen.Model.Elements;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.IO.CodeGen.Model.TopLevel;

/// <summary>
/// <c><![CDATA[<xsd:complexType/>]]></c> that has <c><![CDATA[<xsd:choice>]]></c> as an element.
/// The type is inside <c><![CDATA[<xsd:schema/>]]></c>.
/// <example>
/// <code><![CDATA[
/// <xsd:complexType name="CT_Tables">
///   <xsd:choice minOccurs="1" maxOccurs="unbounded">
///     <xsd:element name="m" type="CT_TableMissing"/>
///     <xsd:element name="s" type="CT_XStringElement"/>
///   </xsd:choice>
///   <xsd:attribute name="count" use="optional" type="xsd:unsignedInt"/>
/// </xsd:complexType>
/// ]]></code>
/// </example>
/// </summary>
public class ComplexTypeChoice : ComplexType, INode
{
    public required Choice Choice { get; init; }

    public T Accept<T>(IXsdVisitor<T> visitor)
    {
        return visitor.Visit(this);
    }

    internal override List<Variable> GenerateParseMethod(CodeBuilder code, string namespaceField)
    {
        var choicesCount = DetermineChoicesCount();

        if (choicesCount == ElementsCount.ZeroToOne)
        {
            // The problem in 0..1 is what to do when nothing is selected. The lister approach doesn't really detect that
            // The best choice for 0..1 is a variable for each choice and pass all possible choices to the hook.

            // Create a variable declarations, one variable for each choice. The values will be passed to the hook.
            var variables = new List<Variable>();
            foreach (var child in Choice.Children)
            {
                var element = (ElementType)child;
                if (code.TryGetComplexType(element.TypeName, out var csType))
                {
                    csType += '?';
                    code.WriteIndent().Append(csType).Append(" ").AppendVariable(element.Name).Append(" = null;").EndLine();
                    variables.Add(new Variable(csType, element.Name));
                }
            }

            var isFirst = true;
            foreach (var child in Choice.Children)
            {
                var element = (ElementType)child;
                code.WriteIndent().Append(!isFirst ? "else " : "").Append($"if (_reader.TryOpen(\"{element.Name}\", {namespaceField}))").EndLine();
                code.OpenBrace();
                code.WriteIndent();
                if (code.TryGetComplexType(element.TypeName, out _))
                    code.AppendVariable(element.Name).Append(" = ");

                code.Append($"Parse{code.NormalizeCt(element.TypeName)}(\"{element.Name}\");").EndLine();
                code.CloseBrace();
                isFirst = false;
            }

            code.AddLine($"_reader.Close(elementName, {namespaceField});");
            return variables;
        }

        if (choicesCount == ElementsCount.OneToMany)
        {
            code.AddLine("do");
            code.OpenBrace();
            var isFirst = true;
            foreach (var child in Choice.Children)
            {
                var element = (ElementType)child;
                var joiner = isFirst ? string.Empty : "else ";
                isFirst = false;

                code.AddLine($"{joiner}if (_reader.TryOpen(\"{element.Name}\", {namespaceField}))");
                code.OpenBrace();
                code.AddLine($"Parse{code.NormalizeCt(element.TypeName)}(\"{element.Name}\");");
                code.CloseBrace();
            }

            code.AddLine("else");
            code.OpenBrace();
            code.AddLine("throw PartStructureException.ExpectedChoiceElementNotFound(_reader);");
            code.CloseBrace();
            code.CloseBrace();
            code.AddLine($"while (!_reader.TryClose(elementName, {namespaceField}));");
        }
        else
        {
            throw new NotImplementedException("Choice element count range is not implemented.");
        }

        return [];
    }

    private ElementsCount DetermineChoicesCount()
    {
        // OOXML XSD is not very consistent with how it defines choices, so normalize
        // the choice to few selected patterns we can implement. Minimum of patterns
        // means simpler and more consistent hooks.
        var min = Choice.Occurrences.Min ?? 1;
        var max = Choice.Occurrences.Max ?? 1;

        var allChoicesSame = Choice.Children.All(x => x is ElementType) &&
                             Choice.Children.Cast<ElementType>().Select(x => x.Occurrences.Elements).Distinct().Count() == 1;

        ElementsCount? choicesElements = allChoicesSame ? Choice.Children.Cast<ElementType>().First().Occurrences.Elements : null;

        // This is pretty ugly, but technically valid XSD. Select one choice from choices
        // that are all optional... Used for CT_Fill and few others.
        if (min == 1 && max == 1 && choicesElements == ElementsCount.ZeroToOne)
        {
            return ElementsCount.ZeroToOne;
        }

        if (min == 1 && max == int.MaxValue && choicesElements == ElementsCount.OneToOne)
        {
            return ElementsCount.OneToMany;
        }

        throw new NotImplementedException($"Unknown code pattern for choice {Name}");
    }
}
