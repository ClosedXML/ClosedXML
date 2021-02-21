// Keep this file CodeMaid organised and cleaned
using ClosedXML.Utils;
using DocumentFormat.OpenXml.Packaging;
using System;

namespace ClosedXML.Excel
{
    public class LoadOptions
    {
        /// <summary>
        /// An factory implementation of <see cref="RelationshipErrorHandler" /> to remove malformed hyperlinks. />
        /// </summary>
        public static Func<OpenXmlPackage, RelationshipErrorHandler> RemoveMalformedHyperlinksRelationshipErrorHandlerFactory { get; } = p => new RemoveMalformedHyperlinksRelationshipErrorHandler(p);

        public XLEventTracking EventTracking { get; set; } = XLEventTracking.Enabled;
        public Boolean RecalculateAllFormulas { get; set; } = false;

        /// <summary>
        /// Gets or sets a delegate that is used to create a handler to rewrite relationships that are malformed. On platforms after .NET 4.5, <see cref="Uri"/> parsing will fail on malformed strings.
        /// </summary>
        public Func<OpenXmlPackage, RelationshipErrorHandler> RelationshipErrorHandlerFactory { get; set; }

        internal OpenSettings ToOpenSettings()
        {
            var settings = new OpenSettings();

            if (this.RelationshipErrorHandlerFactory != null)
            {
                settings.RelationshipErrorHandlerFactory = this.RelationshipErrorHandlerFactory;
            }

            return settings;
        }
    }
}
