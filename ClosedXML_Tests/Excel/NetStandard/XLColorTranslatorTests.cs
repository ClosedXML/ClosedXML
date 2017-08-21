#if _NETSTANDARD_

using ClosedXML.Excel;
using ClosedXML.NetStandard;
using NUnit.Framework;
using System.Linq;
using System.Drawing;

namespace ClosedXML_Tests.Excel.NetStandard
{
    /// <summary>
    ///     Summary description for UnitTest1
    /// </summary>
    [TestFixture]
    public class XLColorTranslatorTests
    {


        [Test]
        public void CanResolveFromHtmlColor()
        {
            Color color;
            color = XLColorTranslator.FromHtml("#FF000000");
            Assert.AreEqual(255, color.A);
            Assert.AreEqual(0, color.R);
            Assert.AreEqual(0, color.G);
            Assert.AreEqual(0, color.B);

            color = XLColorTranslator.FromHtml("#8899AABB");
            Assert.AreEqual(136, color.A);
            Assert.AreEqual(153, color.R);
            Assert.AreEqual(170, color.G);
            Assert.AreEqual(187, color.B);

            color = XLColorTranslator.FromHtml("#99AABB");
            Assert.AreEqual(255, color.A);
            Assert.AreEqual(153, color.R);
            Assert.AreEqual(170, color.G);
            Assert.AreEqual(187, color.B);
        }
    }
}

#endif