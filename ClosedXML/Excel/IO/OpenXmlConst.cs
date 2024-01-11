using System;

namespace ClosedXML.Excel.IO
{
    /// <summary>
    /// Constants used across writers.
    /// </summary>
    internal static class OpenXmlConst
    {
        public const string Main2006SsNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        public const string XmMain2006 = "http://schemas.microsoft.com/office/excel/2006/main";

        public const string X14Main2009SsNs = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

        public const string X14Ac2009SsNs = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";

        public const string Xml1998Ns = "http://www.w3.org/XML/1998/namespace";

        public const string MarkupCompatibilityNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";

        public const string RelationshipsNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        public const string RevisionNs = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";

        /// <summary>
        /// Valid and shorter than normal true.
        /// </summary>
        public static readonly String TrueValue = "1";

        /// <summary>
        /// Valid and shorter than normal false.
        /// </summary>
        public static readonly String FalseValue = "0";
    }
}
