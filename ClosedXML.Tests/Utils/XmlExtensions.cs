#nullable enable
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace ClosedXML.Tests;

internal static class XmlExtensions
{
    private static readonly XAttributeEqualityComparer AttributeComparer = new();

    /// <summary>
    /// Are two XML elements semantically same? Ignores semantically irrelevant differences between
    /// compared elements.
    /// </summary>
    public static bool SemanticallyEqual(this XElement? lhs, XElement? rhs)
    {
        if (lhs is null || rhs is null)
            return lhs is null && rhs is null;

        if (lhs.Name != rhs.Name)
            return false;

        // Namespace declaration is an attribute, but should be ignored for purposes of comparison
        // because they can be declared anywhere, can be declared multiple times and yet
        // semantically be same.
        var lhsAttr = lhs.Attributes().Where(a => !a.IsNamespaceDeclaration).OrderBy(a => a.Name.LocalName).ThenBy(a => a.Name.Namespace);
        var rhsAttr = rhs.Attributes().Where(a => !a.IsNamespaceDeclaration).OrderBy(a => a.Name.LocalName).ThenBy(a => a.Name.Namespace);

        if (!lhsAttr.SequenceEqual(rhsAttr, AttributeComparer))
            return false;

        var lhsNode = lhs.FirstNode;
        var rhsNode = rhs.FirstNode;
        while (lhsNode is not null && rhsNode is not null)
        {
            if (lhsNode.GetType() != rhsNode.GetType())
                return false;

            if (lhsNode is XElement lhsElement && rhsNode is XElement rhsElement)
            {
                if (!lhsElement.SemanticallyEqual(rhsElement))
                    return false;
            }
            else if (lhsNode is XText lhsText && rhsNode is XText rhsText)
            {
                if (lhsText.Value != rhsText.Value)
                    return false;
            }
            else
            {
                throw new NotSupportedException();
            }

            lhsNode = lhsNode.NextNode;
            rhsNode = rhsNode.NextNode;
        }

        return lhsNode is null && rhsNode is null;
    }

    private class XAttributeEqualityComparer : IEqualityComparer<XAttribute>
    {
        public bool Equals(XAttribute? lhs, XAttribute? rhs)
        {
            if (lhs is null || rhs is null)
                return lhs is null && rhs is null;

            return lhs.Name == rhs.Name && lhs.Value == rhs.Value;
        }

        public int GetHashCode(XAttribute obj)
        {
            return obj.GetHashCode();
        }
    }
}
