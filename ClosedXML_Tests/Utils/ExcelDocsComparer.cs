using System;
using System.IO;
using System.IO.Packaging;

namespace ClosedXML_Tests
{
    internal static class ExcelDocsComparer
    {
        public static bool Compare(string left, string right, bool stripColumnWidths, out string message)
        {
            using (FileStream leftStream = File.OpenRead(left))
            {
                using (FileStream rightStream = File.OpenRead(right))
                {
                    return Compare(leftStream, rightStream, stripColumnWidths, out message);
                }
            }
        }

        public static bool Compare(Stream left, Stream right, bool stripColumnWidths, out string message)
        {
            using (Package leftPackage = Package.Open(left))
            {
                using (Package rightPackage = Package.Open(right))
                {
                    return PackageHelper.Compare(leftPackage, rightPackage, false, ExcludeMethod, stripColumnWidths, out message);
                }
            }
        }

        private static bool ExcludeMethod(Uri uri)
        {
            //Exclude service data
            if (uri.OriginalString.EndsWith(".rels") ||
                uri.OriginalString.EndsWith(".psmdcp"))
            {
                return true;
            }
            return false;
        }
    }
}