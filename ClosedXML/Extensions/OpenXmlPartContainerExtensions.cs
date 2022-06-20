// Keep this file CodeMaid organised and cleaned
using DocumentFormat.OpenXml.Packaging;
using System.Linq;

namespace ClosedXML.Extensions
{
    internal static class OpenXmlPartContainerExtensions
    {
        public static bool HasPartWithId(this OpenXmlPartContainer container, string relId)
        {
            return container.Parts.Any(p => p.RelationshipId.Equals(relId));
        }
    }
}
