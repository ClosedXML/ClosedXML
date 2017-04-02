using ClosedXML.Utils;
using NUnit.Framework;

namespace ClosedXML_Tests.Excel
{
    [TestFixture]
    public class XmlEncoderTest
    {
        [Test]
        public void TestControlChars()
        {
            Assert.AreEqual("_x0001_ _x0002_ _x0003_ _x0004_", XmlEncoder.EncodeString("\u0001 \u0002 \u0003 \u0004"));
            Assert.AreEqual("_x0005_ _x0006_ _x0007_ _x0008_", XmlEncoder.EncodeString("\u0005 \u0006 \u0007 \u0008"));
            Assert.AreEqual("\u0001 \u0002 \u0003 \u0004", XmlEncoder.DecodeString("_x0001_ _x0002_ _x0003_ _x0004_"));
            Assert.AreEqual("\u0005 \u0006 \u0007 \u0008", XmlEncoder.DecodeString("_x0005_ _x0006_ _x0007_ _x0008_"));
        }

        [Test]
        public void TestIsXmlChar()
        {
            Assert.AreEqual(false, XmlEncoder.IsXmlChar('\u0001'));
            Assert.AreEqual(false, XmlEncoder.IsXmlChar('\u0005'));
            Assert.AreEqual(false, XmlEncoder.IsXmlChar('\u0007'));
            Assert.AreEqual(false, XmlEncoder.IsXmlChar('\u0008'));
            Assert.AreEqual(true, XmlEncoder.IsXmlChar('J'));
            Assert.AreEqual(true, XmlEncoder.IsXmlChar('+'));
            Assert.AreEqual(true, XmlEncoder.IsXmlChar('S'));
            Assert.AreEqual(true, XmlEncoder.IsXmlChar('4'));
            Assert.AreEqual(true, XmlEncoder.IsXmlChar('!'));
            Assert.AreEqual(true, XmlEncoder.IsXmlChar('$'));
        }
    }
}
