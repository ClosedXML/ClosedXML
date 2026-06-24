using NUnit.Framework.Constraints;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace ClosedXML.Tests;

/// <summary>
/// Compare an element in an <see cref="XDocument"/> with the supplied XML.
/// </summary>
internal class MatchesXmlConstraint(string xml) : Constraint
{
    public override string Description => $"XML should semantically match {xml}.";

    public override ConstraintResult ApplyTo<TActual>(TActual actual)
    {
        if (actual is not IEnumerable<XElement> elements)
            return new ConstraintResult(this, actual, ConstraintStatus.Failure);

        var element = elements.Single();
        var expected = XDocument.Load(new StringReader(xml));
        var xmlEqual = element.SemanticallyEqual(expected.Root);
        return new ConstraintResult(this, element, xmlEqual ? ConstraintStatus.Success : ConstraintStatus.Error);
    }
}
