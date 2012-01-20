using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace ClosedXML_Tests
{
    
    
    /// <summary>
    ///This is a test class for ExcelHelperTest and is intended
    ///to contain all ExcelHelperTest Unit Tests
    ///</summary>
    [TestClass()]
    public class ExcelHelperTest
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
        ///A test for IsValidColumn
        ///</summary>
        [TestMethod()]
        public void IsValidColumnTest()
        {
            Assert.AreEqual(false, ExcelHelper.IsValidColumn(""));
            Assert.AreEqual(false, ExcelHelper.IsValidColumn("1"));
            Assert.AreEqual(false, ExcelHelper.IsValidColumn("A1"));
            Assert.AreEqual(false, ExcelHelper.IsValidColumn("AA1"));
            Assert.AreEqual(true, ExcelHelper.IsValidColumn("A"));
            Assert.AreEqual(true, ExcelHelper.IsValidColumn("AA"));
            Assert.AreEqual(true, ExcelHelper.IsValidColumn("AAA"));
            Assert.AreEqual(true, ExcelHelper.IsValidColumn("Z"));
            Assert.AreEqual(true, ExcelHelper.IsValidColumn("ZZ"));
            Assert.AreEqual(true, ExcelHelper.IsValidColumn("XFD"));
            Assert.AreEqual(false, ExcelHelper.IsValidColumn("ZAA"));
            Assert.AreEqual(false, ExcelHelper.IsValidColumn("XZA"));
            Assert.AreEqual(false, ExcelHelper.IsValidColumn("XFZ"));
        }
    }
}

