using System.IO;
using ClosedXML_Examples;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClosedXML_Tests
{
    [TestClass]
    public class ExcelDocsComparerTests
    {
        [TestMethod]
        public void CheckEqual()
        {
             string left = ExampleHelper.GetTempFilePath("left.xlsx");
             string right = ExampleHelper.GetTempFilePath("right.xlsx");
             try
             {
                 new BasicTable().Create(left);
                 new BasicTable().Create(right);
                 string message;
                 Assert.IsTrue(ExcelDocsComparer.Compare(left, right, out message));
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

        [TestMethod]
        public void CheckNonEqual()
        {
            string left = ExampleHelper.GetTempFilePath("left.xlsx");
            string right = ExampleHelper.GetTempFilePath("right.xlsx");
            try
            {
                new BasicTable().Create(left);
                new HelloWorld().Create(right);

                string message;
                Assert.IsFalse(ExcelDocsComparer.Compare(left, right, out message));
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