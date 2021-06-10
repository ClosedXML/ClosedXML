using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;

namespace ClosedXML.Utils
{
    public class RemoveMalformedHyperlinksRelationshipErrorHandler : RelationshipErrorHandler
    {
        private readonly Dictionary<string, List<string>> _errors;
        private readonly OpenXmlPackage _package;

        public RemoveMalformedHyperlinksRelationshipErrorHandler(OpenXmlPackage package)
        {
            _package = package;
            _errors = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        }

        public override void OnPackageLoaded()
        {
            foreach (var part in _package.GetAllParts())
            {
                if (_errors.TryGetValue(part.Uri.OriginalString, out var ids))
                {
                    foreach (var id in ids)
                    {
                        part.DeleteReferenceRelationship(id);

                        if (part is WorksheetPart && part.RootElement is Worksheet ws)
                        {
                            foreach (var h in ws.Descendants<Hyperlink>())
                            {
                                var parent = h.Parent;

                                if (h.Id == id)
                                {
                                    h.Remove();
                                }

                                if (!parent.HasChildren)
                                {
                                    parent.Remove();
                                }
                            }
                        }
                    }
                }
            }
        }

        public override string Rewrite(Uri partUri, string id, string uri)
        {
            var key = partUri.OriginalString
                .Replace("_rels/", string.Empty)
                .Replace(".rels", string.Empty);

            if (!_errors.ContainsKey(key))
            {
                _errors.Add(key, new List<string>());
            }

            _errors[key].Add(id);

            return "http://error";
        }
    }
}
