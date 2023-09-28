using System;
using System.Collections.ObjectModel;
using System.Globalization;
using ClosedXML.Excel;
using ClosedXML.Excel.IO;
using DocumentFormat.OpenXml;

namespace ClosedXML.Extensions
{
    internal static class OpenXmlPartReaderExtensions
    {
        internal static bool IsStartElement(this OpenXmlPartReader reader, string localName)
        {
            return reader.LocalName == localName && reader.NamespaceUri == OpenXmlConst.Main2006SsNs && reader.IsStartElement;
        }

        internal static void MoveAhead(this OpenXmlPartReader reader)
        {
            if (!reader.Read())
                throw new InvalidOperationException("Unexpected end of stream.");
        }

        internal static string? GetAttribute(this ReadOnlyCollection<OpenXmlAttribute> attributes, string name)
        {
            // Don't use foreach, performance critical
            var length = attributes.Count;
            for (var i = 0; i < length; ++i)
            {
                var attr = attributes[i];
                if (attr.LocalName == name && string.IsNullOrEmpty(attr.NamespaceUri))
                    return attr.Value;
            }

            return null;
        }

        internal static string? GetAttribute(this ReadOnlyCollection<OpenXmlAttribute> attributes, string name, string namespaceUri)
        {
            // Don't use foreach, performance critical
            var length = attributes.Count;
            for (var i = 0; i < length; ++i)
            {
                var attr = attributes[i];
                if (attr.LocalName == name && attr.NamespaceUri == namespaceUri)
                    return attr.Value;
            }

            return null;
        }

        internal static bool GetBoolAttribute(this ReadOnlyCollection<OpenXmlAttribute> attributes, string name, bool defaultValue)
        {
            var attribute = attributes.GetAttribute(name);
            return ParseBool(attribute, defaultValue);
        }

        internal static int? GetIntAttribute(this ReadOnlyCollection<OpenXmlAttribute> attributes, string name)
        {
            var attribute = attributes.GetAttribute(name);
            if (!string.IsNullOrEmpty(attribute))
                return int.Parse(attribute);

            return null;
        }

        internal static uint? GetUintAttribute(this ReadOnlyCollection<OpenXmlAttribute> attributes, string name)
        {
            var attribute = attributes.GetAttribute(name);
            if (!string.IsNullOrEmpty(attribute))
                return uint.Parse(attribute);

            return null;
        }

        internal static double? GetDoubleAttribute(this ReadOnlyCollection<OpenXmlAttribute> attributes, string name, string namespaceUri)
        {
            var attribute = attributes.GetAttribute(name, namespaceUri);
            if (!string.IsNullOrEmpty(attribute))
                return double.Parse(attribute, NumberStyles.Float, XLHelper.ParseCulture);

            return null;
        }

        internal static double? GetDoubleAttribute(this ReadOnlyCollection<OpenXmlAttribute> attributes, string name)
        {
            var attribute = attributes.GetAttribute(name);
            if (!string.IsNullOrEmpty(attribute))
                return double.Parse(attribute, NumberStyles.Float, XLHelper.ParseCulture);

            return null;
        }

        /// <summary>
        /// Get value of attribute with type <c>ST_CellRef</c>.
        /// </summary>
        internal static XLSheetPoint? GetCellRefAttribute(this ReadOnlyCollection<OpenXmlAttribute> attributes, string name)
        {
            var attribute = attributes.GetAttribute(name);
            if (!string.IsNullOrEmpty(attribute))
                return XLSheetPoint.Parse(attribute);

            return null;
        }

        /// <summary>
        /// Get value of attribute with type <c>ST_Ref</c>.
        /// </summary>
        internal static XLSheetRange? GetRefAttribute(this ReadOnlyCollection<OpenXmlAttribute> attributes, string name)
        {
            var attribute = attributes.GetAttribute(name);
            if (!string.IsNullOrEmpty(attribute))
                return XLSheetRange.Parse(attribute);

            return null;
        }

        private static bool ParseBool(string? input, bool defaultValue)
        {
            if (string.IsNullOrEmpty(input))
                return defaultValue;

            var isTrue = input == "1" || string.Equals("true", input, StringComparison.OrdinalIgnoreCase);
            if (isTrue)
                return true;

            var isFalse = input == "0" || string.Equals("false", input, StringComparison.OrdinalIgnoreCase);
            if (isFalse)
                return false;

            throw new FormatException($"Unable to parse '{input}' to bool.");
        }
    }
}
