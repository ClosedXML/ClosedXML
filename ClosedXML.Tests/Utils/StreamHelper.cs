using MoreLinq;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using System.Xml.XPath;

namespace ClosedXML.Tests
{
    /// <summary>
    ///     Help methods for work with streams
    /// </summary>
    public static class StreamHelper
    {
        private static readonly XName colTagName = XName.Get("col", @"http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        private static readonly XName widthAttrName = XName.Get("width");
        private static readonly IEnumerable<(string PartUri, string NodesXPath)> uriSpecificIgnores = new List<(string PartUri, string NodesXPath)>()
        {
            // Remove dcterms elements, e.g. created, modified
            ("/docProps/core.xml", "//*[namespace-uri() = 'http://purl.org/dc/terms/']")
        };

        /// <summary>
        ///     Convert stream to byte array
        /// </summary>
        /// <param name="pStream">Stream</param>
        /// <returns>Byte array</returns>
        public static byte[] StreamToArray(Stream pStream)
        {
            long iLength = pStream.Length;
            var bytes = new byte[iLength];
            for (int i = 0; i < iLength; i++)
            {
                bytes[i] = (byte)pStream.ReadByte();
            }
            pStream.Close();
            return bytes;
        }

        /// <summary>
        ///     Convert byte array to stream
        /// </summary>
        /// <param name="pBynaryArray">Byte array</param>
        /// <param name="pStream">Open stream</param>
        /// <returns></returns>
        public static Stream ArrayToStreamAppend(byte[] pBynaryArray, Stream pStream)
        {
            #region Check params

            if (ReferenceEquals(pBynaryArray, null))
            {
                throw new ArgumentNullException("pBynaryArray");
            }
            if (ReferenceEquals(pStream, null))
            {
                throw new ArgumentNullException("pStream");
            }
            if (!pStream.CanWrite)
            {
                throw new ArgumentException("Can't write to stream", "pStream");
            }

            #endregion Check params

            foreach (byte b in pBynaryArray)
            {
                pStream.WriteByte(b);
            }
            return pStream;
        }

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
        /// <param name="one"></param>
        /// <param name="other"></param>
        /// /// <param name="stripColumnWidths"></param>
        /// <returns></returns>
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


            var leftString = new StreamReader(left.Stream).ReadToEnd().RemoveIgnoredParts(ignoreGuids: true);
            var rightString = new StreamReader(right.Stream).ReadToEnd().RemoveIgnoredParts(ignoreGuids: true);

            var isXmlContent = left.ContentType.EndsWith("+xml");
            if (!isXmlContent)
            {
                return leftString == rightString;
            }

            var leftXml = XDocument.Parse(leftString);
            var rightXml = XDocument.Parse(rightString);

            var toleranceInPercent = stripColumnWidths ? decimal.MaxValue : 0.03m;
            var leftWidths = GetWidths(leftXml);
            var rightWidths = GetWidths(rightXml);

            var areWeightsEqual = AreWidthsEqual(leftWidths, rightWidths, toleranceInPercent);
            if (!areWeightsEqual) return false;

            CopyWeights(rightWidths, leftWidths);

            RemoveIgnoredParts(leftXml, partUri);
            RemoveIgnoredParts(rightXml, partUri);

            return XNode.DeepEquals(leftXml, rightXml);
        }

        private static List<(XAttribute Attr, decimal Width)> GetWidths(XDocument document)
        {
            return document.Descendants(colTagName)
                .Select(tag => tag.Attribute(widthAttrName))
                .Where(attr => !(attr is null))
                .Select(attr => (attr, decimal.Parse(attr.Value, CultureInfo.InvariantCulture)))
                .ToList();
        }

        private static bool AreWidthsEqual(List<(XAttribute Attr, decimal Width)> leftWidths, List<(XAttribute Attr, decimal Width)> rightWidths, decimal tolerance)
        {
            // This checks values and count, position of width attributes is checked when comparing whole document
            if (leftWidths.Count != rightWidths.Count)
                return false;

            return leftWidths.Zip(rightWidths, (First, Second) => (First, Second)).All(t =>
            {
                var delta = Math.Abs(t.Second.Width - t.First.Width);
                return Math.Max(delta / t.First.Width, delta / t.Second.Width) <= tolerance;
            });
        }

        private static void CopyWeights(List<(XAttribute Attr, decimal Width)> fromWidths, List<(XAttribute Attr, decimal Width)> toWidths)
        {
            fromWidths.Zip(toWidths, (From, To) => (From, To)).ForEach(t =>
            {
                t.To.Attr.SetValue(t.From.Attr.Value);
            });
        }

        private static void RemoveIgnoredParts(XDocument document, Uri partUri)
        {
            foreach (var (_, NodesXPath) in uriSpecificIgnores.Where(p => p.PartUri.Equals(partUri.OriginalString)))
                foreach (var node in document.XPathSelectElements(NodesXPath).ToList())
                    node.Remove();
        }

        private static string RemoveIgnoredParts(this string s, Boolean ignoreGuids)
        {
            // Collapse empty xml elements
            s = emptyXmlElementRegex.Replace(s, "<$1 />");

            if (ignoreGuids)
                s = RemoveGuids(s);

            return s;
        }


        private static Regex emptyXmlElementRegex = new Regex(@"<([\w:]+)><\/\1>", RegexOptions.Compiled);

        private static Regex guidRegex = new Regex(@"{[0-9A-Fa-f]{8}-([0-9A-Fa-f]{4}-){3}[0-9A-Fa-f]{12}}", RegexOptions.Compiled | RegexOptions.Multiline);

        private static String RemoveGuids(String s)
        {
            return guidRegex.Replace(s, delegate (Match m)
            {
                return string.Empty;
            });
        }
    }
}
