using System.IO;
using ClosedXML_Examples;
using NUnit.Framework;

namespace ClosedXML_Tests
{
    [TestFixture]
    public class ExcelDocsComparerTests
    {
        [Test]
        public void CheckEqual()
        {
            string left = ExampleHelper.GetTempFilePath("left.xlsx");
            string right = ExampleHelper.GetTempFilePath("right.xlsx");
            try
            {
                new BasicTable().Create(left);
                new BasicTable().Create(right);
                string message;
#if _NETFRAMEWORK_
                Assert.IsTrue(ExcelDocsComparer.Compare(left, right, TestHelper.IsRunningOnUnix, out message));
#else
                Assert.IsTrue(ExcelDocsComparer.Compare(left, right, true, out message));
#endif                
            }
            finally
            {
                if (File.Exists(left))
                {
                    File.Delete(left);
                }
                if (File.Exists(right))
                {
                    File.Delete(right);
                }
            }
        }

        [Test]
        public void CheckNonEqual()
        {
            string left = ExampleHelper.GetTempFilePath("left.xlsx");
            string right = ExampleHelper.GetTempFilePath("right.xlsx");
            try
            {
                new BasicTable().Create(left);
                new HelloWorld().Create(right);

                string message;
#if _NETFRAMEWORK_                
                Assert.IsFalse(ExcelDocsComparer.Compare(left, right, TestHelper.IsRunningOnUnix, out message));
#else                
                Assert.IsFalse(ExcelDocsComparer.Compare(left, right, true, out message));
#endif                
            }
            finally
            {
                if (File.Exists(left))
                {
                    File.Delete(left);
                }
                if (File.Exists(right))
                {
                    File.Delete(right);
                }
            }
        }
    }
}