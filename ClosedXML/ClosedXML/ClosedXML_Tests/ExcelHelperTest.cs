using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace ClosedXML_Tests
{
    ///<summary>
    ///	This is a test class for XLHelperTest and is intended
    ///	to contain all XLHelperTest Unit Tests
    ///</summary>
    [TestClass]
    public class XLHelperTest
    {
        ///<summary>
        ///	Gets or sets the test context which provides
        ///	information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext { get; set; }

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

        ///<summary>
        ///	A test for IsValidColumn
        ///</summary>
        [TestMethod]
        public void IsValidColumnTest()
        {
            Assert.AreEqual(false, XLHelper.IsValidColumn(""));
            Assert.AreEqual(false, XLHelper.IsValidColumn("1"));
            Assert.AreEqual(false, XLHelper.IsValidColumn("A1"));
            Assert.AreEqual(false, XLHelper.IsValidColumn("AA1"));
            Assert.AreEqual(true, XLHelper.IsValidColumn("A"));
            Assert.AreEqual(true, XLHelper.IsValidColumn("AA"));
            Assert.AreEqual(true, XLHelper.IsValidColumn("AAA"));
            Assert.AreEqual(true, XLHelper.IsValidColumn("Z"));
            Assert.AreEqual(true, XLHelper.IsValidColumn("ZZ"));
            Assert.AreEqual(true, XLHelper.IsValidColumn("XFD"));
            Assert.AreEqual(false, XLHelper.IsValidColumn("ZAA"));
            Assert.AreEqual(false, XLHelper.IsValidColumn("XZA"));
            Assert.AreEqual(false, XLHelper.IsValidColumn("XFZ"));
        }

        [TestMethod]
        public void ReplaceRelative1()
        {
            var result = XLHelper.ReplaceRelative("A1", 2, "B");
            Assert.AreEqual("B2", result);
        }

        [TestMethod]
        public void ReplaceRelative2()
        {
            var result = XLHelper.ReplaceRelative("$A1", 2, "B");
            Assert.AreEqual("$A2", result);
        }

        [TestMethod]
        public void ReplaceRelative3()
        {
            var result = XLHelper.ReplaceRelative("A$1", 2, "B");
            Assert.AreEqual("B$1", result);
        }

        [TestMethod]
        public void ReplaceRelative4()
        {
            var result = XLHelper.ReplaceRelative("$A$1", 2, "B");
            Assert.AreEqual("$A$1", result);
        }

        [TestMethod]
        public void ReplaceRelative5()
        {
            var result = XLHelper.ReplaceRelative("1:1", 2, "B");
            Assert.AreEqual("2:2", result);
        }

        [TestMethod]
        public void ReplaceRelative6()
        {
            var result = XLHelper.ReplaceRelative("$1:1", 2, "B");
            Assert.AreEqual("$1:2", result);
        }

        [TestMethod]
        public void ReplaceRelative7()
        {
            var result = XLHelper.ReplaceRelative("1:$1", 2, "B");
            Assert.AreEqual("2:$1", result);
        }

        [TestMethod]
        public void ReplaceRelative8()
        {
            var result = XLHelper.ReplaceRelative("$1:$1", 2, "B");
            Assert.AreEqual("$1:$1", result);
        }

        [TestMethod]
        public void ReplaceRelative9()
        {
            var result = XLHelper.ReplaceRelative("A:A", 2, "B");
            Assert.AreEqual("B:B", result);
        }

        [TestMethod]
        public void ReplaceRelativeA()
        {
            var result = XLHelper.ReplaceRelative("$A:A", 2, "B");
            Assert.AreEqual("$A:B", result);
        }

        [TestMethod]
        public void ReplaceRelativeB()
        {
            var result = XLHelper.ReplaceRelative("A:$A", 2, "B");
            Assert.AreEqual("B:$A", result);
        }

        [TestMethod]
        public void ReplaceRelativeC()
        {
            var result = XLHelper.ReplaceRelative("$A:$A", 2, "B");
            Assert.AreEqual("$A:$A", result);
        }
    }
}