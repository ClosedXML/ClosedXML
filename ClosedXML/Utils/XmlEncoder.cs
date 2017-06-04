using System.Text;
using System.Xml;

namespace ClosedXML.Utils
{
    public static class XmlEncoder
    {
        /// <summary>
        /// Checks if a character is not allowed to the XML Spec http://www.w3.org/TR/REC-xml/#charsets
        /// </summary>
        /// <param name="ch">Input Character</param>
        /// <returns>Returns false if the character is invalid according to the XML specification, and will not be
        /// escaped by an XmlWriter.</returns>
        public static bool IsXmlChar(char ch)
        {
            return (((ch >= 0x0020 && ch <= 0xD7FF) ||
                      (ch >= 0xE000 && ch <= 0xFFFD) ||
                      ch == 0x0009 || ch == 0x000A ||
                      ch == 0x000D));
        }

        public static string EncodeString(string encodeStr)
        {
            if (encodeStr == null) return null;

            var newString = new StringBuilder();

            foreach (var ch in encodeStr)
            {
                if (IsXmlChar(ch)) //this method is new in .NET 4
                {
                    newString.Append(ch);
                }
                else
                {
                    newString.Append(XmlConvert.EncodeName(ch.ToString()));
                }
            }
            return newString.ToString();
        }

        public static string DecodeString(string decodeStr)
        {
            return XmlConvert.DecodeName(decodeStr);
        }
    }
}
