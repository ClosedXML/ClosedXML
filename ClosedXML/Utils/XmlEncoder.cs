using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace ClosedXML.Utils
{
    public static class XmlEncoder
    {
        private static readonly Regex xHHHHRegex = new Regex("_(x[\\dA-Fa-f]{4})_", RegexOptions.Compiled);

        public static string EncodeString(string encodeStr)
        {
            if (encodeStr == null) return null;

            encodeStr = xHHHHRegex.Replace(encodeStr, "_x005F_$1_");

            var sb = new StringBuilder(encodeStr.Length);

            foreach (var ch in encodeStr)
            {
                if (XmlConvert.IsXmlChar(ch))
                {
                    sb.Append(ch);
                }
                else
                {
                    sb.Append(XmlConvert.EncodeName(ch.ToString()));
                }
            }

            return sb.ToString();
        }

        public static string DecodeString(string decodeStr)
        {
            return XmlConvert.DecodeName(decodeStr);
        }
    }
}
