using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace ClosedXML.Tests
{
    /// <summary>
    ///     Help methods for work with streams
    /// </summary>
    public static class StreamHelper
    {
        private static readonly XName colTagName = XName.Get("col", @"http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        private static readonly XName widthAttrName = XName.Get("width");

        private static readonly IEnumerable<(string PartSubstring, XName NodeName)> ignoredNodes = new List<(string PartSubstring, XName NodeName)>
        {
            ("/docProps/core.xml", XName.Get("created", @"http://purl.org/dc/terms/")),
            ("/docProps/core.xml", XName.Get("modified", @"http://purl.org/dc/terms/")),
            ("sheet", XName.Get("id", @"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"))
        };

        private static readonly IEnumerable<(string PartSubstring, XName NodeName, XName AttrName)> ignoredAttributes = new List<(string PartSubstring, XName NodeName, XName AttrName)>
        {
            ("sheet", XName.Get("cfRule", @"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"), XName.Get("id"))
        };

        public static void StreamToStreamAppend(Stream streamIn, Stream streamToWrite)
        {
            StreamToStreamAppend(streamIn, streamToWrite, 0);
        }

        public static void StreamToStreamAppend(Stream streamIn, Stream streamToWrite, long dataLength)
        {
            #region Check params

            if (ReferenceEquals(streamIn, null))
            {
                throw new ArgumentNullException("streamIn");
            }
            if (ReferenceEquals(streamToWrite, null))
            {
                throw new ArgumentNullException("streamToWrite");
            }
            if (!streamIn.CanRead)
            {
                throw new ArgumentException("Can't read from stream", "streamIn");
            }
            if (!streamToWrite.CanWrite)
            {
                throw new ArgumentException("Can't write to stream", "streamToWrite");
            }

            #endregion Check params

            var buf = new byte[512];
            long length;
            if (dataLength == 0)
            {
                length = streamIn.Length - streamIn.Position;
            }
            else
            {
                length = dataLength;
            }
            long rest = length;
            while (rest > 0)
            {
                int len1 = streamIn.Read(buf, 0, rest >= 512 ? 512 : (int)rest);
                streamToWrite.Write(buf, 0, len1);
                rest -= len1;
            }
        }

        /// <summary>
        ///     Compare two streams by converting them to strings and comparing the strings
        /// </summary>
        public static bool Compare((string ContentType, Stream Stream) left, (string ContentType, Stream Stream) right, Uri partUri, bool stripColumnWidths)
        {
            #region Check

            if (left.ContentType == null || left.Stream == null)
            {
                throw new ArgumentNullException(nameof(left));
            }
            if (right.ContentType == null || right.Stream == null)
            {
                throw new ArgumentNullException(nameof(right));
            }
            if (left.Stream.Position != 0)
            {
                throw new ArgumentException("Must be in position 0", nameof(left));
            }
            if (right.Stream.Position != 0)
            {
                throw new ArgumentException("Must be in position 0", nameof(right));
            }

            if (left.ContentType != right.ContentType)
            {
                throw new ArgumentException("Different content types.");
            }

            #endregion Check

            var leftString = new StreamReader(left.Stream).ReadToEnd();
            var rightString = new StreamReader(right.Stream).ReadToEnd();

            var isXmlContent = left.ContentType.EndsWith("+xml");
            if (!isXmlContent)
                return leftString == rightString;

            var leftXml = XDocument.Parse(leftString);
            RemoveIgnoredParts(leftXml, partUri, stripColumnWidths);
            Normalize(leftXml);
            var rightXml = XDocument.Parse(rightString);
            RemoveIgnoredParts(rightXml, partUri, stripColumnWidths);
            Normalize(rightXml);

            return XNode.DeepEquals(leftXml, rightXml);
        }

        private static void RemoveIgnoredParts(XDocument document, Uri partUri, Boolean stripColumnWidths)
        {
            foreach (var ignoredNode in ignoredNodes.Where(i => partUri.OriginalString.Contains(i.PartSubstring)))
                document.Descendants(ignoredNode.NodeName).Remove();

            foreach (var ignoredAttr in ignoredAttributes.Where(i => partUri.OriginalString.Contains(i.PartSubstring)))
                document.Descendants(ignoredAttr.NodeName).Attributes(ignoredAttr.AttrName).Remove();

            if (stripColumnWidths)
                document.Descendants(colTagName).Attributes(widthAttrName).Remove();
        }

        private static void Normalize(XDocument document)
        {
            // Turn empty elements into self closing ones.
            foreach (var emptyElement in document.Descendants().Where(e => !e.IsEmpty && !e.Nodes().Any()))
                emptyElement.RemoveNodes();

            foreach (var element in document.Descendants().Where(e => e.Attributes().Any()))
            {
                var attrs = element.Attributes().OrderBy(a => a.Name.LocalName).ToList();
                element.ReplaceAttributes(attrs);
            }
        }
    }
}
