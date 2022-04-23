using ClosedXML.Examples;
using NUnit.Framework;
using System.IO;

namespace ClosedXML.Tests
{
    [TestFixture]
    public class ExcelDocsComparerTests
    {
        [Test]
        public void CheckEqual()
        {
            var left = ExampleHelper.GetTempFilePath("left.xlsx");
            var right = ExampleHelper.GetTempFilePath("right.xlsx");
            try
            {
                new BasicTable().Create(left);
                new BasicTable().Create(right);
                Assert.IsTrue(ExcelDocsComparer.Compare(left, right, out var message));
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
            var left = ExampleHelper.GetTempFilePath("left.xlsx");
            var right = ExampleHelper.GetTempFilePath("right.xlsx");
            try
            {
                new BasicTable().Create(left);
                new HelloWorld().Create(right);

                Assert.IsFalse(ExcelDocsComparer.Compare(left, right, out var message));
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
