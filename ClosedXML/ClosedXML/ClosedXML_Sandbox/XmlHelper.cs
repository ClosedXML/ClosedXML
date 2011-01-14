using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using System.IO;
using System.Xml;

namespace ClosedXML_Sandbox
{
    public class XmlHelper
    {
        //Creates an object from an XML string.

        /// <summary>
        /// Creates an object from an XML string.
        /// </summary>
        /// <typeparam name="T">The type of object to be constructed.</typeparam>
        /// <param name="xml">The XML string to convert to the object.</param>
        /// <returns>An object of type T.</returns>
        public static T GetObjectFromXml<T>(string xml)
        {
            T retVal;
            Type objType = typeof(T);
            var ser = new XmlSerializer(objType);
            using (var stringReader = new StringReader(xml))
            {
                retVal = (T)ser.Deserialize(stringReader);
                stringReader.Close();
            }
            return retVal;
        }

        /// <summary>
        /// Creates an XML string from an object.
        /// </summary>
        /// <typeparam name="T">The type of object process.</typeparam>
        /// <param name="obj">The object from which to extract the XML.</param>
        /// <returns>An XML string.</returns>
        public static String GetXmlFromObject<T>(T obj, XmlSerializerNamespaces xmlSerializerNamespaces = null)
        {
            String retVal;
            Type objType = typeof(T);
            var ser = new XmlSerializer(objType);
            //this will remove the namespace from the xml. Necessary per RealEC.
            //XmlSerializerNamespaces xmlnsEmpty = new XmlSerializerNamespaces();
            //xmlnsEmpty.Add(String.Empty, String.Empty);

            using (var memStream = new MemoryStream())
            {

                ser.Serialize(memStream, obj, xmlSerializerNamespaces);

                retVal = Encoding.UTF8.GetString(memStream.GetBuffer()).Replace("\0", "");
                memStream.Close();
            }

            return retVal;
        }

        public static String ConvertSpecialChars(String xmlData)
        {
            return
                xmlData
                .Replace("&lt;", "<")
                .Replace("#60;", "<")
                .Replace("&gt;", ">")
                .Replace("&#62;", ">")
                .Replace("&quot;", "\"")
                .Replace("&#39;", "\"")
                .Replace("&apos;", "'")
                .Replace("&#34;", "'")
                .Replace("&amp;", "&")
                .Replace("#38;", "&");
        }
    }
}
