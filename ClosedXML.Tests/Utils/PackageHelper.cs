using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Net.Mime;
using System.Text;
using System.Xml.Serialization;

namespace ClosedXML.Tests.Utils
{
    public static class PackageHelper
    {
        public static void WriteXmlPart(Package package, Uri uri, object content, XmlSerializer serializer)
        {
            if (package.PartExists(uri))
            {
                package.DeletePart(uri);
            }
            var part = package.CreatePart(uri, MediaTypeNames.Text.Xml, CompressionOption.Fast);
            using var stream = part.GetStream();
            serializer.Serialize(stream, content);
        }

        public static object ReadXmlPart(Package package, Uri uri, XmlSerializer serializer)
        {
            if (!package.PartExists(uri))
            {
                throw new ApplicationException(string.Format("Package part '{0}' doesn't exists!", uri.OriginalString));
            }
            var part = package.GetPart(uri);
            using var stream = part.GetStream();
            return serializer.Deserialize(stream);
        }

        public static void WriteBinaryPart(Package package, Uri uri, Stream content)
        {
            if (package.PartExists(uri))
            {
                package.DeletePart(uri);
            }
            var part = package.CreatePart(uri, MediaTypeNames.Application.Octet, CompressionOption.Fast);
            using var stream = part.GetStream();
            StreamHelper.StreamToStreamAppend(content, stream);
        }

        /// <summary>
        ///     Returns part's stream
        /// </summary>
        /// <param name="package"></param>
        /// <param name="uri"></param>
        /// <returns></returns>
        public static Stream ReadBinaryPart(Package package, Uri uri)
        {
            if (!package.PartExists(uri))
            {
                throw new ApplicationException("Package part doesn't exists!");
            }
            var part = package.GetPart(uri);
            return part.GetStream();
        }

        public static void CopyPart(Uri uri, Package source, Package dest)
        {
            CopyPart(uri, source, dest, true);
        }

        public static void CopyPart(Uri uri, Package source, Package dest, bool overwrite)
        {
            #region Check

            if (uri is null)
            {
                throw new ArgumentNullException("uri");
            }
            if (source is null)
            {
                throw new ArgumentNullException("source");
            }
            if (dest is null)
            {
                throw new ArgumentNullException("dest");
            }

            #endregion Check

            if (dest.PartExists(uri))
            {
                if (!overwrite)
                {
                    throw new ArgumentException("Specified part already exists", "uri");
                }
                dest.DeletePart(uri);
            }

            var sourcePart = source.GetPart(uri);
            var destPart = dest.CreatePart(uri, sourcePart.ContentType, sourcePart.CompressionOption);

            using var sourceStream = sourcePart.GetStream();
            using var destStream = destPart.GetStream();
            StreamHelper.StreamToStreamAppend(sourceStream, destStream);
        }

        public static void WritePart<T>(Package package, PackagePartDescriptor descriptor, T content,
            Action<Stream, T> serializeAction)
        {
            #region Check

            if (package is null)
            {
                throw new ArgumentNullException("package");
            }
            if (descriptor is null)
            {
                throw new ArgumentNullException("descriptor");
            }
            if (serializeAction is null)
            {
                throw new ArgumentNullException("serializeAction");
            }

            #endregion Check

            if (package.PartExists(descriptor.Uri))
            {
                package.DeletePart(descriptor.Uri);
            }
            var part = package.CreatePart(descriptor.Uri, descriptor.ContentType, descriptor.CompressOption);
            using var stream = part.GetStream();
            serializeAction(stream, content);
        }

        public static void WritePart(Package package, PackagePartDescriptor descriptor, Action<Stream> serializeAction)
        {
            #region Check

            if (package is null)
            {
                throw new ArgumentNullException("package");
            }
            if (descriptor is null)
            {
                throw new ArgumentNullException("descriptor");
            }
            if (serializeAction is null)
            {
                throw new ArgumentNullException("serializeAction");
            }

            #endregion Check

            if (package.PartExists(descriptor.Uri))
            {
                package.DeletePart(descriptor.Uri);
            }
            var part = package.CreatePart(descriptor.Uri, descriptor.ContentType, descriptor.CompressOption);
            using var stream = part.GetStream();
            serializeAction(stream);
        }

        public static T ReadPart<T>(Package package, Uri uri, Func<Stream, T> deserializeFunc)
        {
            #region Check

            if (package is null)
            {
                throw new ArgumentNullException("package");
            }
            if (uri is null)
            {
                throw new ArgumentNullException("uri");
            }
            if (deserializeFunc is null)
            {
                throw new ArgumentNullException("deserializeFunc");
            }

            #endregion Check

            if (!package.PartExists(uri))
            {
                throw new ApplicationException(string.Format("Package part '{0}' doesn't exists!", uri.OriginalString));
            }
            var part = package.GetPart(uri);
            using var stream = part.GetStream();
            return deserializeFunc(stream);
        }

        public static void ReadPart(Package package, Uri uri, Action<Stream> deserializeAction)
        {
            #region Check

            if (package is null)
            {
                throw new ArgumentNullException("package");
            }
            if (uri is null)
            {
                throw new ArgumentNullException("uri");
            }
            if (deserializeAction is null)
            {
                throw new ArgumentNullException("deserializeAction");
            }

            #endregion Check

            if (!package.PartExists(uri))
            {
                throw new ApplicationException(string.Format("Package part '{0}' doesn't exists!", uri.OriginalString));
            }
            var part = package.GetPart(uri);
            using var stream = part.GetStream();
            deserializeAction(stream);
        }

        public static bool TryReadPart(Package package, Uri uri, Action<Stream> deserializeAction)
        {
            #region Check

            if (package is null)
            {
                throw new ArgumentNullException("package");
            }
            if (uri is null)
            {
                throw new ArgumentNullException("uri");
            }
            if (deserializeAction is null)
            {
                throw new ArgumentNullException("deserializeAction");
            }

            #endregion Check

            if (!package.PartExists(uri))
            {
                return false;
            }
            var part = package.GetPart(uri);
            using (var stream = part.GetStream())
            {
                deserializeAction(stream);
            }
            return true;
        }

        /// <summary>
        ///     Compare to packages by parts like streams
        /// </summary>
        /// <param name="left"></param>
        /// <param name="right"></param>
        /// <param name="compareToFirstDifference"></param>
        /// <param name="excludeMethod"></param>
        /// <param name="message"></param>
        /// <returns></returns>
        public static bool Compare(Package left, Package right, bool compareToFirstDifference, out string message)
        {
            return Compare(left, right, compareToFirstDifference, null, out message, false);
        }

        /// <summary>
        ///     Compare to packages by parts like streams
        /// </summary>
        /// <param name="left"></param>
        /// <param name="right"></param>
        /// <param name="compareToFirstDifference"></param>
        /// <param name="excludeMethod"></param>
        /// <param name="message"></param>
        /// <returns></returns>
        public static bool Compare(Package left, Package right, bool compareToFirstDifference,
            Func<Uri, bool> excludeMethod, out string message, bool ignoreColumnFormat)
        {
            #region Check

            if (left == null)
            {
                throw new ArgumentNullException("left");
            }
            if (right == null)
            {
                throw new ArgumentNullException("right");
            }

            #endregion Check

            excludeMethod = excludeMethod ?? (uri => false);
            var leftParts = left.GetParts();
            var rightParts = right.GetParts();

            var pairs = new Dictionary<Uri, PartPair>();
            foreach (var part in leftParts)
            {
                if (excludeMethod(part.Uri))
                {
                    continue;
                }
                pairs.Add(part.Uri, new PartPair(part.Uri, CompareStatus.OnlyOnLeft));
            }
            foreach (var part in rightParts)
            {
                if (excludeMethod(part.Uri))
                {
                    continue;
                }
                if (pairs.TryGetValue(part.Uri, out var pair))
                {
                    pair.Status = CompareStatus.Equal;
                }
                else
                {
                    pairs.Add(part.Uri, new PartPair(part.Uri, CompareStatus.OnlyOnRight));
                }
            }

            if (compareToFirstDifference && pairs.Any(pair => pair.Value.Status != CompareStatus.Equal))
            {
                return AgregateCompareResult(out message, pairs);
            }

            foreach (var pair in pairs.Values)
            {
                if (pair.Status != CompareStatus.Equal)
                {
                    continue;
                }
                var leftPart = left.GetPart(pair.Uri);
                var rightPart = right.GetPart(pair.Uri);
                using var leftPackagePartStream = leftPart.GetStream(FileMode.Open, FileAccess.Read);
                using var rightPackagePartStream = rightPart.GetStream(FileMode.Open, FileAccess.Read);
                using var leftMemoryStream = new MemoryStream();
                using var rightMemoryStream = new MemoryStream();
                leftPackagePartStream.CopyTo(leftMemoryStream);
                rightPackagePartStream.CopyTo(rightMemoryStream);

                leftMemoryStream.Seek(0, SeekOrigin.Begin);
                rightMemoryStream.Seek(0, SeekOrigin.Begin);

                var stripColumnWidthsFromSheet = ignoreColumnFormat &&
                    leftPart.ContentType == @"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" &&
                    rightPart.ContentType == @"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

                var tuple1 = new Tuple<Uri, Stream>(pair.Uri, leftMemoryStream);
                var tuple2 = new Tuple<Uri, Stream>(pair.Uri, rightMemoryStream);

                if (!StreamHelper.Compare(tuple1, tuple2, stripColumnWidthsFromSheet))
                {
                    pair.Status = CompareStatus.NonEqual;
                    if (compareToFirstDifference)
                    {
                        return AgregateCompareResult(out message, pairs);
                    }
                }
            }

            var sortedPairs = pairs.Values.ToList();
            sortedPairs.Sort((one, other) => one.Uri.OriginalString.CompareTo(other.Uri.OriginalString));
            var sbuilder = new StringBuilder();
            foreach (var pair in sortedPairs)
            {
                if (pair.Status == CompareStatus.Equal)
                {
                    continue;
                }
                sbuilder.AppendFormat("{0} :{1}", pair.Uri, pair.Status);
                sbuilder.AppendLine();
            }
            message = sbuilder.ToString();
            return message.Length == 0;
        }

        private static bool AgregateCompareResult(out string message, Dictionary<Uri, PartPair> pairs)
        {
            var sortedPairs = pairs.Values.ToList();
            sortedPairs.Sort((one, other) => one.Uri.OriginalString.CompareTo(other.Uri.OriginalString));
            var sbuilder = new StringBuilder();
            foreach (var pair in sortedPairs)
            {
                if (pair.Status == CompareStatus.Equal)
                {
                    continue;
                }
                sbuilder.AppendFormat("{0} :{1}", pair.Uri, pair.Status);
                sbuilder.AppendLine();
            }
            message = sbuilder.ToString();
            return message.Length == 0;
        }

        public sealed class PackagePartDescriptor
        {
            [DebuggerBrowsable(DebuggerBrowsableState.Never)]
            private readonly CompressionOption _compressOption;

            [DebuggerBrowsable(DebuggerBrowsableState.Never)]
            private readonly string _contentType;

            [DebuggerBrowsable(DebuggerBrowsableState.Never)]
            private readonly Uri _uri;

            /// <summary>
            ///     Instance constructor
            /// </summary>
            /// <param name="uri">Part uri</param>
            /// <param name="contentType">Content type from <see cref="MediaTypeNames" /></param>
            /// <param name="compressOption"></param>
            public PackagePartDescriptor(Uri uri, string contentType, CompressionOption compressOption)
            {
                if (uri is null)
                {
                    throw new ArgumentNullException("uri");
                }
                if (string.IsNullOrEmpty(contentType))
                {
                    throw new ArgumentNullException("contentType");
                }

                _uri = uri;
                _contentType = contentType;
                _compressOption = compressOption;
            }

            public Uri Uri
            {
                [DebuggerStepThrough]
                get { return _uri; }
            }

            public string ContentType
            {
                [DebuggerStepThrough]
                get { return _contentType; }
            }

            public CompressionOption CompressOption
            {
                [DebuggerStepThrough]
                get { return _compressOption; }
            }

            public override string ToString()
            {
                return string.Format("Uri:{0} ContentType: {1}, Compression: {2}", _uri, _contentType, _compressOption);
            }
        }

        #region Nested type: CompareStatus

        private enum CompareStatus
        {
            OnlyOnLeft,
            OnlyOnRight,
            Equal,
            NonEqual
        }

        #endregion Nested type: CompareStatus

        private sealed class PartPair
        {
            [DebuggerBrowsable(DebuggerBrowsableState.Never)]
            private readonly Uri _uri;

            [DebuggerBrowsable(DebuggerBrowsableState.Never)]
            private CompareStatus _status;

            public PartPair(Uri uri, CompareStatus status)
            {
                _uri = uri;
                _status = status;
            }

            public Uri Uri
            {
                [DebuggerStepThrough]
                get { return _uri; }
            }

            public CompareStatus Status
            {
                [DebuggerStepThrough]
                get { return _status; }
                [DebuggerStepThrough]
                set { _status = value; }
            }
        }

        //--
    }
}