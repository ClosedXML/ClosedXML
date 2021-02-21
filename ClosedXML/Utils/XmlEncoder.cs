using NeoSmart.Unicode;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace ClosedXML.Utils
{
    public static class XmlEncoder
    {
        private static readonly Regex xHHHHRegex = new Regex("_(x[\\dA-Fa-f]{4})_", RegexOptions.Compiled);
        private static readonly Regex Uppercase_X_HHHHRegex = new Regex("_(X[\\dA-Fa-f]{4})_", RegexOptions.Compiled);

        public static string EncodeString(string encodeStr)
        {
            if (encodeStr == null) return null;

            encodeStr = xHHHHRegex.Replace(encodeStr, "_x005F_$1_");

            var sb = new StringBuilder(encodeStr.Length);

            foreach (var letter in encodeStr.Letters())
            {

                if (Emoji.IsEmoji(letter))
                {
                    sb.Append(letter);
                    continue;
                }

                foreach (var ch in letter)
                {
                    if (char.IsSymbol(ch) || XmlConvert.IsXmlChar(ch))
                    {
                        sb.Append(letter);
                    }
                    else
                    {
                        sb.Append(XmlConvert.EncodeName(ch.ToString()));
                    }
                }
            }

            return sb.ToString();
        }

        public static string DecodeString(string decodeStr)
        {
            if (string.IsNullOrEmpty(decodeStr)) return string.Empty;

            // Strings "escaped" with _X (capital X) should not be treated as escaped
            // Example: _Xceed_Something
            // https://github.com/ClosedXML/ClosedXML/issues/1154
            decodeStr = Uppercase_X_HHHHRegex.Replace(decodeStr, "_x005F_$1_");

            return XmlConvert.DecodeName(decodeStr);
        }
    }
}
