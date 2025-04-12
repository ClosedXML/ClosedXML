using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace ClosedXML.Utils
{
    internal static class XmlEncoder
    {
        private static readonly Regex xHHHHRegex = new Regex("_(x[\\dA-Fa-f]{4})_", RegexOptions.Compiled);

        public static string EncodeString(string encodeStr)
        {
            encodeStr = xHHHHRegex.Replace(encodeStr, "_x005F_$1_");

            var sb = new StringBuilder(encodeStr.Length);
            var len = encodeStr.Length;
            for (var i = 0; i < len; ++i)
            {
                var currentChar = encodeStr[i];
                if (XmlConvert.IsXmlChar(currentChar))
                {
                    sb.Append(currentChar);
                }
                else if (i + 1 < len && XmlConvert.IsXmlSurrogatePair(encodeStr[i + 1], currentChar))
                {
                    sb.Append(currentChar);
                    sb.Append(encodeStr[++i]);
                }
                else
                {
                    sb.Append(XmlConvert.EncodeName(currentChar.ToString()));
                }
            }

            return sb.ToString();
        }
    }
}
