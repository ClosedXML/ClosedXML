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
            string left = ExampleHelper.GetTempFilePath("left.xlsx");
            string right = ExampleHelper.GetTempFilePath("right.xlsx");
            try
            {
                new BasicTable().Create(left);
                new BasicTable().Create(right);
                Assert.That(ExcelDocsComparer.Compare(left, right, out string message), Is.True);
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

                Assert.That(ExcelDocsComparer.Compare(left, right, out string message), Is.False);
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
