using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ClosedXML_Tests
{
    /// <summary>
    ///     Help methods for work with streams
    /// </summary>
    public static class StreamHelper
    {
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
        public static bool Compare(Tuple<Uri, Stream> tuple1, Tuple<Uri, Stream> tuple2, bool stripColumnWidths)
        {
            #region Check

            if (tuple1 == null || tuple1.Item1 == null || tuple1.Item2 == null)
            {
                throw new ArgumentNullException("one");
            }
            if (tuple2 == null || tuple2.Item1 == null || tuple2.Item2 == null)
            {
                throw new ArgumentNullException("other");
            }
            if (tuple1.Item2.Position != 0)
            {
                throw new ArgumentException("Must be in position 0", "one");
            }
            if (tuple2.Item2.Position != 0)
            {
                throw new ArgumentException("Must be in position 0", "other");
            }

            #endregion Check

            var stringOne = new StreamReader(tuple1.Item2).ReadToEnd().RemoveIgnoredParts(tuple1.Item1, stripColumnWidths, ignoreGuids: true);
            var stringOther = new StreamReader(tuple2.Item2).ReadToEnd().RemoveIgnoredParts(tuple2.Item1, stripColumnWidths, ignoreGuids: true);
            return stringOne == stringOther;
        }

        private static string RemoveIgnoredParts(this string s, Uri uri, Boolean ignoreColumnWidths, Boolean ignoreGuids)
        {
            foreach (var pair in uriSpecificIgnores.Where(p => p.Key.Equals(uri.OriginalString)))
                s = pair.Value.Replace(s, "");

            // Collapse empty xml elements
            s = emptyXmlElementRegex.Replace(s, "<$1 />");

            if (ignoreColumnWidths)
                s = RemoveColumnWidths(s);

            if (ignoreGuids)
                s = RemoveGuids(s);

            return s;
        }

        private static IEnumerable<KeyValuePair<string, Regex>> uriSpecificIgnores = new List<KeyValuePair<string, Regex>>()
        {
            // Remove dcterms elements
            new KeyValuePair<string, Regex>("/docProps/core.xml", new Regex(@"<dcterms:(\w+).*?<\/dcterms:\1>", RegexOptions.Compiled))
        };

        private static Regex emptyXmlElementRegex = new Regex(@"<([\w:]+)><\/\1>", RegexOptions.Compiled);
        private static Regex columnRegex = new Regex("<x:col.*?width=\"\\d+(\\.\\d+)?\".*?\\/>", RegexOptions.Compiled);
        private static Regex widthRegex = new Regex("width=\"\\d+(\\.\\d+)?\"\\s+", RegexOptions.Compiled);

        private static String RemoveColumnWidths(String s)
        {
            var replacements = new Dictionary<String, String>();

            foreach (var m in columnRegex.Matches(s).OfType<Match>())
            {
                var original = m.Groups[0].Value;
                var replacement = widthRegex.Replace(original, "");
                replacements.Add(original, replacement);
            }

            foreach (var r in replacements)
            {
                s = s.Replace(r.Key, r.Value);
            }
            return s;
        }

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
