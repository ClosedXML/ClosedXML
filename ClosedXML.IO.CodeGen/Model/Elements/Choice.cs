using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.IO.CodeGen.Model.Elements;

/// <summary>
/// <c><![CDATA[<xsd:choice>]]></c> inside <c><![CDATA[<xsd:complexType>]]></c>.
/// </summary>
public class Choice : IElementGroup
{
    public required List<IElementGroup> Children { get; init; } = [];

    public required Occurrences Occurrences { get; init; }

    public T Accept<T>(IXsdVisitor<T> visitor)
    {
        return visitor.Visit(this);
    }

    internal List<Variable> GenerateParseContent(CodeBuilder code, string namespaceField)
    {
        var choicesCount = DetermineChoicesCount();

        List<Variable> variables;
        switch (choicesCount)
        {
            case ElementsCount.ZeroToOne:
                {
                    variables = GenerateParseContent(choicesCount, code, namespaceField);
                    code.AddLine($"_reader.Close(elementName, {namespaceField});");
                    break;
                }
            case ElementsCount.OneToMany:
                {
                    code.AddLine("do");
                    code.OpenBrace();
                    variables = GenerateParseContent(choicesCount, code, namespaceField);
                    code.CloseBrace();
                    code.AddLine($"while (!_reader.TryClose(elementName, {namespaceField}));");
                    break;
                }
            default:
                throw new NotImplementedException();
        }

        return variables;
    }

    internal List<Variable> GenerateParseContent(ElementsCount choicesCount, CodeBuilder code, string namespaceField)
    {
        if (choicesCount is ElementsCount.ZeroToOne)
        {
            // The problem in 0..1 is what to do when nothing is selected. The lister approach doesn't really detect that
            // The best choice for 0..1 is a variable for each choice and pass all possible choices to the hook.

            // Create a variable declarations, one variable for each choice. The values will be passed to the hook.
            var variables = DeclareChildrenVariables(code);

            var isFirst = true;
            foreach (var child in Children)
            {
                var element = (ElementType)child;
                code.WriteIndent().Append(!isFirst ? "else " : "").Append($"if (_reader.TryOpen(\"{element.Name}\", {namespaceField}))").EndLine();
                code.OpenBrace();
                code.WriteIndent();
                if (code.TryGetCsType(element.TypeName, out _))
                    code.AppendVariable(element.Name).Append(" = ");

                code.Append($"Parse{code.NormalizeCt(element.TypeName)}(\"{element.Name}\");").EndLine();
                code.CloseBrace();
                isFirst = false;
            }

            return variables;
        }

        if (choicesCount == ElementsCount.OneToMany)
        {
            var isFirst = true;
            foreach (var child in Children)
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
        }
        else
        {
            throw new NotImplementedException("Choice element count range is not implemented.");
        }

        return [];
    }

    private List<Variable> DeclareChildrenVariables(CodeBuilder code)
    {
        var variables = new List<Variable>();
        foreach (var child in Children)
        {
            var element = (ElementType)child;
            if (code.TryGetCsType(element.TypeName, out var csType))
            {
                csType += '?';
                code.WriteIndent().Append(csType).Append(" ").AppendVariable(element.Name).Append(" = null;").EndLine();
                variables.Add(new Variable(csType, element.Name));
            }
        }

        return variables;
    }

    internal ElementsCount DetermineChoicesCount()
    {
        // OOXML XSD is not very consistent with how it defines choices, so normalize
        // the choice to few selected patterns we can implement. Minimum of patterns
        // means simpler and more consistent hooks.
        var min = Occurrences.Min ?? 1;
        var max = Occurrences.Max ?? 1;

        var allChoicesSame = Children.All(x => x is ElementType) &&
                             Children.Cast<ElementType>().Select(x => x.Occurrences.Elements).Distinct().Count() == 1;

        ElementsCount? choicesElements = allChoicesSame ? Children.Cast<ElementType>().First().Occurrences.Elements : null;

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

        if (min == 1 && max == 1 && choicesElements == ElementsCount.OneToOne)
        {
            return ElementsCount.OneToOne;
        }

        throw new NotImplementedException($"Unknown code pattern for a choice.");
    }
}
