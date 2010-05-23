using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace ClosedXML_Test
{
    
    /// <summary>
    ///This is a test class for XLWorksheetsTest and is intended
    ///to contain all XLWorksheetsTest Unit Tests
    ///</summary>
    [TestClass()]
    public class XLWorksheetsTest
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
        ///A test for XLWorksheets Constructor
        ///</summary>
        [TestMethod()]
        public void XLWorksheets_Constructor_Test()
        {
            var wbExample = new XLWorkbook(@"c:\Example.xlsx");
            XLWorksheets target = new XLWorksheets(wbExample);

        }

        [TestMethod()]
        public void XLWorksheets_Count0_Add1_Add2_Delete1_Count1_Test()
        {
            var wbExample = new XLWorkbook(@"c:\Example.xlsx");
            XLWorksheets target = new XLWorksheets(wbExample);

            Assert.AreEqual(target.Count, 0U);

            target.Add("Sheet1");

            Assert.AreEqual(target.Count, 1U);

            target.Add("Sheet2");

            Assert.AreEqual(target.Count, 2U);

            target.Delete("Sheet2");

            Assert.AreEqual(target.Count, 1U);

        }
    }
}
