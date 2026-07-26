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

    internal List<Variable> GenerateParseContent(ParsletName parsletName, ElementsCount choicesCount, CodeBuilder code, bool throwOnFail)
    {
        // Check whether containing top-level element returns a value.
        // The output type `choice`, `choice?`, a list, or choice0..
        if (!code.TryGetCsType(parsletName, out var choiceCsType))
            choiceCsType = null;

        if (choicesCount == ElementsCount.ZeroToOne)
        {
            // Declare result variable
            const string resultVariableName = "choice";
            if (choiceCsType is not null)
            {
                code.AddLine($"{choiceCsType}? {resultVariableName};");
            }

            AddOneChoice(parsletName, resultVariableName, code);

            // Add final branch. No child matched for 0-1 choice, so it's fine.
            if (choiceCsType is not null)
            {
                code.AddLine("else");
                code.OpenBrace();
                code.AddLine($"{resultVariableName} = default;");
                code.CloseBrace();
                return [new Variable(choiceCsType, resultVariableName)];
            }

            return [];
        }

        if (choicesCount == ElementsCount.OneToOne)
        {
            // Declare result variable
            const string resultVariableName = "choice";
            if (choiceCsType is not null)
            {
                code.AddLine($"{choiceCsType} {resultVariableName};");
            }

            AddOneChoice(parsletName, resultVariableName, code);

            // Add final branch. No child matched for 1-1 choice, so it's a fail.
            code.AddLine("else");
            code.OpenBrace();
            if (throwOnFail)
            {
                code.AddLine("throw PartStructureException.ExpectedChoiceElementNotFound(_reader);");
            }
            else
            {
                var returnValue = choiceCsType is not null ? $"Xpr.Fail<{choiceCsType}>()" : "Xpr.Fail()";
                code.AddLine($"return {returnValue};");
            }

            code.CloseBrace();

            if (choiceCsType is not null)
            {
                return [new Variable(choiceCsType, resultVariableName)];
            }

            return [];
        }

        if (choicesCount is ElementsCount.ZeroToMany or ElementsCount.OneToMany)
        {
            // Make a list variable
            // For each matched choice call hook.
            var minChoices = choicesCount == ElementsCount.OneToMany ? 1 : 0;
            code.AddLine($"// Choice with cardinality {minChoices}-n");

            // Declare a output variable 
            const string resultVariableName = "choiceList";
            if (choiceCsType is not null)
            {
                var itemType = code.GetCsItemType(parsletName);
                code.AddLine($"var {resultVariableName} = new List<{itemType}>();");
            }

            var checkCount = minChoices > 0;
            if (checkCount)
            {
                code.AddLine("var choiceCount = 0;");
            }

            code.AddLine("while (true)");
            code.OpenBrace();

            // A variable for one iteration
            const string iterationVariableName = "choice";
            if (choiceCsType is not null)
            {
                var itemType = code.GetCsItemType(parsletName);
                code.AddLine($"{itemType} {iterationVariableName};");
            }

            AddOneChoice(parsletName, iterationVariableName, code);

            code.AddLine("else");
            code.OpenBrace();

            // No choice element was matched => Break out of a cycle, choice sequence has ended
            code.AddLine("break;");
            code.CloseBrace();

            if (choiceCsType is not null)
            {
                code.AddLine($"{resultVariableName}.Add({iterationVariableName});");
            }

            if (checkCount)
            {
                code.AddLine("choiceCount++;");
            }

            // End of while
            code.CloseBrace();

            if (checkCount)
            {
                code.AddLine("if(choiceCount == 0)");
                code.OpenBrace();
                code.AddLine("throw PartStructureException.IncorrectElementsCount();");
                code.CloseBrace();
            }

            if (choiceCsType is not null)
            {
                return [new Variable(choiceCsType, resultVariableName)];
            }

            return [];
        }

        throw new NotImplementedException();
    }

    private void AddOneChoice(ParsletName parsletName, string resultVariableName, CodeBuilder code)
    {
        // Go over each choice and return first one that parser was able to parse.
        var isFirst = true;
        foreach (var child in Children)
        {
            if (child is ElementType elementChild)
            {
                if (elementChild.Occurrences.ActualMax > 1)
                    throw new NotSupportedException($"Top level element choice {parsletName} needs a custom logic.");

                // For each element, add a hook when a choice is used. It uses a combination of
                // parslet name and element within the choice. That is done, because several have
                // multiple choices with same element type.
                var choiceHookName = parsletName.Value + char.ToUpperInvariant(elementChild.Name[0]) + elementChild.Name[1..];

                // Basically same code, depends whether the the choice returns a value
                if (code.TryGetCsType(elementChild.TypeName, out var elementCsType))
                {
                    // else if (ParseChoiceChild("choiceChildElementName") is { IsSuccess: true } choiceChildResult)
                    //     choice = OnChoiceChild(choiceChildResult.Value);
                    var elementVariable = new Variable(elementCsType, elementChild.Name);
                    code.WriteIndent().Append(isFirst ? string.Empty : "else ").Append("if (").AppendCtParseCall(elementChild.TypeName, elementChild.Name).Append(" is { IsSuccess: true } ").AppendVariable(elementVariable.Name).Append(")").EndLine(); 
                    code.OpenBrace();
                    code.WriteIndent().Append(resultVariableName).Append(" = ").AppendCallHook(choiceHookName, [elementVariable with { Name = elementVariable.Name + ".Value" }]).Append(";").EndLine();
                    code.CloseBrace();
                }
                else
                {
                    // else if (ParseChoiceChild("choiceChildElementName") is { IsSuccess: true })
                    //     // Choice choiceName was successfully parsed
                    code.WriteIndent().Append(isFirst ? string.Empty : "else ").Append("if (").AppendCtParseCall(elementChild.TypeName, elementChild.Name).Append(" is { IsSuccess: true }").Append(")").EndLine();
                    code.OpenBrace();
                    code.AddLine($"// Choice {elementChild.Name} was successfully parsed");
                    code.CloseBrace();
                }
            }
            else
            {
                throw new NotImplementedException();
            }

            isFirst = false;
        }
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

        if (min == 0 && max == 1 && choicesElements == ElementsCount.OneToOne)
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

        if (min == 0 && max == int.MaxValue && choicesElements == ElementsCount.OneToOne)
        {
            return ElementsCount.ZeroToMany;
        }

        throw new NotImplementedException("Unknown code pattern for a choice.");
    }
}
