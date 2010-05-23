using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;

namespace ClosedXML_Test
{
    
    
    /// <summary>
    ///This is a test class for XLWorkbookTest and is intended
    ///to contain all XLWorkbookTest Unit Tests
    ///</summary>
    [TestClass()]
    public class XLWorkbookTest
    {


        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region Additional test attributes
        // 
        //You can use the following additional attributes as you write your tests:
        //
        //Use ClassInitialize to run code before running the first test in the class
        //[ClassInitialize()]
        //public static void MyClassInitialize(TestContext testContext)
        //{
        //}
        //
        //Use ClassCleanup to run code after all tests in a class have run
        //[ClassCleanup()]
        //public static void MyClassCleanup()
        //{
        //}
        //
        //Use TestInitialize to run code before running each test
        //[TestInitialize()]
        //public void MyTestInitialize()
        //{
        //}
        //
        //Use TestCleanup to run code after each test has run
        //[TestCleanup()]
        //public void MyTestCleanup()
        //{
        //}
        //
        #endregion


        /// <summary>
        ///A test for XLWorkbook Constructor
        ///</summary>
        [TestMethod()]
        public void XLWorkbookConstructorTest()
        {
            FileInfo fi = new FileInfo("Test.xlsx");
            XLWorkbook target = new XLWorkbook(fi.FullName);

            Assert.AreEqual(target.Name, fi.Name);
            Assert.AreEqual(target.FullName, fi.FullName);
        }
    }
}
