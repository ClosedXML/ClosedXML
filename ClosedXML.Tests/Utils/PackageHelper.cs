using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Net.Mime;
using System.Text;
using System.Xml.Serialization;

namespace ClosedXML.Tests
{
    public static class PackageHelper
    {
        public static void WriteXmlPart(Package package, Uri uri, object content, XmlSerializer serializer)
        {
            if (package.PartExists(uri))
            {
                package.DeletePart(uri);
            }
            PackagePart part = package.CreatePart(uri, MediaTypeNames.Text.Xml, CompressionOption.Fast);
            using (Stream stream = part.GetStream())
            {
                serializer.Serialize(stream, content);
            }
        }

        public static object ReadXmlPart(Package package, Uri uri, XmlSerializer serializer)
        {
            if (!package.PartExists(uri))
            {
                throw new ApplicationException(string.Format("Package part '{0}' doesn't exists!", uri.OriginalString));
            }
            PackagePart part = package.GetPart(uri);
            using (Stream stream = part.GetStream())
            {
                return serializer.Deserialize(stream);
            }
        }

        public static void WriteBinaryPart(Package package, Uri uri, Stream content)
        {
            if (package.PartExists(uri))
            {
                package.DeletePart(uri);
            }
            PackagePart part = package.CreatePart(uri, MediaTypeNames.Application.Octet, CompressionOption.Fast);
            using (Stream stream = part.GetStream())
            {
                StreamHelper.StreamToStreamAppend(content, stream);
            }
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
            PackagePart part = package.GetPart(uri);
            return part.GetStream();
        }

        public static void CopyPart(Uri uri, Package source, Package dest)
        {
            CopyPart(uri, source, dest, true);
        }

        public static void CopyPart(Uri uri, Package source, Package dest, bool overwrite)
        {
            #region Check

            if (ReferenceEquals(uri, null))
            {
                throw new ArgumentNullException("uri");
            }
            if (ReferenceEquals(source, null))
            {
                throw new ArgumentNullException("source");
            }
            if (ReferenceEquals(dest, null))
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

            PackagePart sourcePart = source.GetPart(uri);
            PackagePart destPart = dest.CreatePart(uri, sourcePart.ContentType, sourcePart.CompressionOption);

            using (Stream sourceStream = sourcePart.GetStream())
            {
                using (Stream destStream = destPart.GetStream())
                {
                    StreamHelper.StreamToStreamAppend(sourceStream, destStream);
                }
            }
        }

        public static void WritePart<T>(Package package, PackagePartDescriptor descriptor, T content,
            Action<Stream, T> serializeAction)
        {
            #region Check

            if (ReferenceEquals(package, null))
            {
                throw new ArgumentNullException("package");
            }
            if (ReferenceEquals(descriptor, null))
            {
                throw new ArgumentNullException("descriptor");
            }
            if (ReferenceEquals(serializeAction, null))
            {
                throw new ArgumentNullException("serializeAction");
            }

            #endregion Check

            if (package.PartExists(descriptor.Uri))
            {
                package.DeletePart(descriptor.Uri);
            }
            PackagePart part = package.CreatePart(descriptor.Uri, descriptor.ContentType, descriptor.CompressOption);
            using (Stream stream = part.GetStream())
            {
                serializeAction(stream, content);
            }
        }

        public static void WritePart(Package package, PackagePartDescriptor descriptor, Action<Stream> serializeAction)
        {
            #region Check

            if (ReferenceEquals(package, null))
            {
                throw new ArgumentNullException("package");
            }
            if (ReferenceEquals(descriptor, null))
            {
                throw new ArgumentNullException("descriptor");
            }
            if (ReferenceEquals(serializeAction, null))
            {
                throw new ArgumentNullException("serializeAction");
            }

            #endregion Check

            if (package.PartExists(descriptor.Uri))
            {
                package.DeletePart(descriptor.Uri);
            }
            PackagePart part = package.CreatePart(descriptor.Uri, descriptor.ContentType, descriptor.CompressOption);
            using (Stream stream = part.GetStream())
            {
                serializeAction(stream);
            }
        }

        public static T ReadPart<T>(Package package, Uri uri, Func<Stream, T> deserializeFunc)
        {
            #region Check

            if (ReferenceEquals(package, null))
            {
                throw new ArgumentNullException("package");
            }
            if (ReferenceEquals(uri, null))
            {
                throw new ArgumentNullException("uri");
            }
            if (ReferenceEquals(deserializeFunc, null))
            {
                throw new ArgumentNullException("deserializeFunc");
            }

            #endregion Check

            if (!package.PartExists(uri))
            {
                throw new ApplicationException(string.Format("Package part '{0}' doesn't exists!", uri.OriginalString));
            }
            PackagePart part = package.GetPart(uri);
            using (Stream stream = part.GetStream())
            {
                return deserializeFunc(stream);
            }
        }

        public static void ReadPart(Package package, Uri uri, Action<Stream> deserializeAction)
        {
            #region Check

            if (ReferenceEquals(package, null))
            {
                throw new ArgumentNullException("package");
            }
            if (ReferenceEquals(uri, null))
            {
                throw new ArgumentNullException("uri");
            }
            if (ReferenceEquals(deserializeAction, null))
            {
                throw new ArgumentNullException("deserializeAction");
            }

            #endregion Check

            if (!package.PartExists(uri))
            {
                throw new ApplicationException(string.Format("Package part '{0}' doesn't exists!", uri.OriginalString));
            }
            PackagePart part = package.GetPart(uri);
            using (Stream stream = part.GetStream())
            {
                deserializeAction(stream);
            }
        }

        public static bool TryReadPart(Package package, Uri uri, Action<Stream> deserializeAction)
        {
            #region Check

            if (ReferenceEquals(package, null))
            {
                throw new ArgumentNullException("package");
            }
            if (ReferenceEquals(uri, null))
            {
                throw new ArgumentNullException("uri");
            }
            if (ReferenceEquals(deserializeAction, null))
            {
                throw new ArgumentNullException("deserializeAction");
            }

            #endregion Check

            if (!package.PartExists(uri))
            {
                return false;
            }
            PackagePart part = package.GetPart(uri);
            using (Stream stream = part.GetStream())
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
            return Compare(left, right, compareToFirstDifference, null, out message);
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
            Func<Uri, bool> excludeMethod, out string message)
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
            PackagePartCollection leftParts = left.GetParts();
            PackagePartCollection rightParts = right.GetParts();

            var pairs = new Dictionary<Uri, PartPair>();
            foreach (PackagePart part in leftParts)
            {
                if (excludeMethod(part.Uri))
                {
                    continue;
                }
                pairs.Add(part.Uri, new PartPair(part.Uri, CompareStatus.OnlyOnLeft));
            }
            foreach (PackagePart part in rightParts)
            {
                if (excludeMethod(part.Uri))
                {
                    continue;
                }
                if (pairs.TryGetValue(part.Uri, out PartPair pair))
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
                goto EXIT;
            }

            foreach (PartPair pair in pairs.Values)
            {
                if (pair.Status != CompareStatus.Equal)
                {
                    continue;
                }
                var leftPart = left.GetPart(pair.Uri);
                var rightPart = right.GetPart(pair.Uri);
                using (Stream leftPackagePartStream = leftPart.GetStream(FileMode.Open, FileAccess.Read))
                using (Stream rightPackagePartStream = rightPart.GetStream(FileMode.Open, FileAccess.Read))
                using (var leftMemoryStream = new MemoryStream())
                using (var rightMemoryStream = new MemoryStream())
                {
                    leftPackagePartStream.CopyTo(leftMemoryStream);
                    rightPackagePartStream.CopyTo(rightMemoryStream);

                    leftMemoryStream.Seek(0, SeekOrigin.Begin);
                    rightMemoryStream.Seek(0, SeekOrigin.Begin);

                    bool stripColumnWidthsFromSheet = TestHelper.StripColumnWidths &&
                        leftPart.ContentType == @"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" &&
                        rightPart.ContentType == @"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

                    var tuple1 = (leftPart.ContentType, Stream: leftMemoryStream);
                    var tuple2 = (rightPart.ContentType, Stream: rightMemoryStream);

                    if (!StreamHelper.Compare(tuple1, tuple2, pair.Uri, stripColumnWidthsFromSheet))
                    {
                        pair.Status = CompareStatus.NonEqual;
                        if (compareToFirstDifference)
                        {
                            goto EXIT;
                        }
                    }
                }
            }

        EXIT:
            List<PartPair> sortedPairs = pairs.Values.ToList();
            sortedPairs.Sort((one, other) => one.Uri.OriginalString.CompareTo(other.Uri.OriginalString));
            var sbuilder = new StringBuilder();
            foreach (PartPair pair in sortedPairs)
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

        #region Nested type: PackagePartDescriptor

        public sealed class PackagePartDescriptor
        {
            #region Private fields

            [DebuggerBrowsable(DebuggerBrowsableState.Never)]
            private readonly CompressionOption _compressOption;

            [DebuggerBrowsable(DebuggerBrowsableState.Never)]
            private readonly string _contentType;

            [DebuggerBrowsable(DebuggerBrowsableState.Never)]
            private readonly Uri _uri;

            #endregion Private fields

            #region Constructor

            /// <summary>
            ///     Instance constructor
            /// </summary>
            /// <param name="uri">Part uri</param>
            /// <param name="contentType">Content type from <see cref="MediaTypeNames" /></param>
            /// <param name="compressOption"></param>
            public PackagePartDescriptor(Uri uri, string contentType, CompressionOption compressOption)
            {
                #region Check

                if (ReferenceEquals(uri, null))
                {
                    throw new ArgumentNullException("uri");
                }
                if (string.IsNullOrEmpty(contentType))
                {
                    throw new ArgumentNullException("contentType");
                }

                #endregion Check

                _uri = uri;
                _contentType = contentType;
                _compressOption = compressOption;
            }

            #endregion Constructor

            #region Public properties

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

            #endregion Public properties

            #region Public methods

            public override string ToString()
            {
                return string.Format("Uri:{0} ContentType: {1}, Compression: {2}", _uri, _contentType, _compressOption);
            }

            #endregion Public methods
        }

        #endregion Nested type: PackagePartDescriptor

        #region Nested type: CompareStatus

        private enum CompareStatus
        {
            OnlyOnLeft,
            OnlyOnRight,
            Equal,
            NonEqual
        }

        #endregion Nested type: CompareStatus

        #region Nested type: PartPair

        private sealed class PartPair
        {
            #region Private fields

            [DebuggerBrowsable(DebuggerBrowsableState.Never)]
            private readonly Uri _uri;

            [DebuggerBrowsable(DebuggerBrowsableState.Never)]
            private CompareStatus _status;

            #endregion Private fields

            #region Constructor

            public PartPair(Uri uri, CompareStatus status)
            {
                _uri = uri;
                _status = status;
            }

            #endregion Constructor

            #region Public properties

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

            #endregion Public properties
        }

        #endregion Nested type: PartPair

        //--
    }
}
