using ClosedXML.Utils;
using NUnit.Framework;
using static ClosedXML.Excel.XLProtectionAlgorithm;

namespace ClosedXML.Tests.Excel.Protection
{
    public class HashAlgorithmTests
    {
        [Test]
        public void TestEmptyPassword()
        {
            Assert.IsEmpty(CryptographicAlgorithms.GetPasswordHash(Algorithm.SHA512, string.Empty));
            Assert.IsEmpty(CryptographicAlgorithms.GetPasswordHash(Algorithm.SimpleHash, string.Empty));
        }

        [Test]
        public void TestSHA512()
        {
            var hash = CryptographicAlgorithms.GetPasswordHash(Algorithm.SHA512, "12345", "aVvPw1DNH3evPqRAd/y3UQ==", 100000);
            Assert.AreEqual("E+qAhyIg/HM0dUrPaENfimFOZp7wlOkJsf/sdG+AGHOA9grOv7VLb1ik2vuYohljI9G36e0ea9wnixCK0MMuyQ==", hash);
        }

        [Test]
        public void TestSimple()
        {
            var hash = CryptographicAlgorithms.GetPasswordHash(Algorithm.SimpleHash, "12345");
            Assert.AreEqual("CA9C", hash);
        }
    }
}
