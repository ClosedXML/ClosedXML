using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace ClosedXML_Test
{
    
    
    /// <summary>
    ///This is a test class for XLAddressTest and is intended
    ///to contain all XLAddressTest Unit Tests
    ///</summary>
    [TestClass()]
    public class XLAddressTest
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
        ///A test for XLAddress Constructor
        ///</summary>
        [TestMethod()]
        public void XLAddress_R1C1Constructor_Test()
        {
            uint row = 4;
            uint column = 4;
            XLAddress target = new XLAddress(row, column);

            Assert.AreEqual(target.Row, 4U);
            Assert.AreEqual(target.Column, 4U);
            Assert.AreEqual(target.ColumnLetter, "D");
        }

        [TestMethod()]
        public void XLAddress_A1Constructor_Test()
        {
            String address = "D4";
            XLAddress target = new XLAddress(address);

            Assert.AreEqual(target.Row, 4U);
            Assert.AreEqual(target.Column, 4U);
            Assert.AreEqual(target.ColumnLetter, "D");
        }

        [TestMethod()]
        public void XLAddress_MixedConstructor_Test()
        {
            uint row = 4; 
            String columnLetter = "D";
            XLAddress target = new XLAddress(row, columnLetter);

            Assert.AreEqual(target.Row, 4U);
            Assert.AreEqual(target.Column, 4U);
            Assert.AreEqual(target.ColumnLetter, "D");
        }

        /// <summary>
        ///A test for GetColumnNumberFromLetter
        ///</summary>
        [TestMethod()]
        public void GetColumnNumberFromLetterTest()
        {
            string column = "OMV";
            uint expected = 10500;
            uint actual;
            actual = XLAddress.GetColumnNumberFromLetter(column);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        ///A test for GetColumnLetterFromNumber
        ///</summary>
        [TestMethod()]
        public void GetColumnLetterFromNumberTest()
        {
            uint column = 10500; 
            string expected = "OMV";
            string actual;
            actual = XLAddress.GetColumnLetterFromNumber(column);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        ///A test for op_Addition
        ///</summary>
        [TestMethod()]
        public void op_AdditionTest()
        {
            XLAddress xlCellAddressLeft = new XLAddress(7, 5); 
            uint right = 1; 
            XLAddress expected = new XLAddress(8,6); 
            XLAddress actual;
            actual = (xlCellAddressLeft + right);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        ///A test for op_Addition
        ///</summary>
        [TestMethod()]
        public void op_AdditionTest1()
        {
            XLAddress xlCellAddressLeft = new XLAddress(7,5); 
            XLAddress xlCellAddressRight = new XLAddress(10, 3);
            XLAddress expected = new XLAddress(17, 8); 
            XLAddress actual;
            actual = (xlCellAddressLeft + xlCellAddressRight);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        ///A test for op_Equality
        ///</summary>
        [TestMethod()]
        public void op_EqualityTest()
        {
            XLAddress xlCellAddressLeft = new XLAddress(1,1); 
            XLAddress xlCellAddressRight = new XLAddress(1,1);
            bool expected = true; 
            bool actual;
            actual = (xlCellAddressLeft == xlCellAddressRight);
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void op_EqualityTest1()
        {
            XLAddress xlCellAddressLeft = new XLAddress(1, 1);
            XLAddress xlCellAddressRight = new XLAddress(1, 2);
            bool expected = false;
            bool actual;
            actual = (xlCellAddressLeft == xlCellAddressRight);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        ///A test for op_GreaterThan
        ///</summary>
        [TestMethod()]
        public void op_GreaterThanTest()
        {
            XLAddress xlCellAddressLeft = new XLAddress(3,3);
            XLAddress xlCellAddressRight = new XLAddress(2,4);
            bool expected = true; 
            bool actual;
            actual = (xlCellAddressLeft > xlCellAddressRight);
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void op_GreaterThanTest1()
        {
            XLAddress xlCellAddressLeft = new XLAddress(3, 3);
            XLAddress xlCellAddressRight = new XLAddress(4, 1);
            bool expected = true;
            bool actual;
            actual = (xlCellAddressLeft > xlCellAddressRight);
            Assert.AreEqual(expected, actual);
        }

        public void op_GreaterThanTest2()
        {
            XLAddress xlCellAddressLeft = new XLAddress(3, 3);
            XLAddress xlCellAddressRight = new XLAddress(3, 4);
            bool expected = false;
            bool actual;
            actual = (xlCellAddressLeft > xlCellAddressRight);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        ///A test for op_GreaterThanOrEqual
        ///</summary>
        [TestMethod()]
        public void op_GreaterThanOrEqualTest()
        {
            XLAddress xlCellAddressLeft = new XLAddress(3,3);
            XLAddress xlCellAddressRight = new XLAddress(2,4);
            bool expected = true; 
            bool actual;
            actual = (xlCellAddressLeft >= xlCellAddressRight);
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void op_GreaterThanOrEqualTest1()
        {
            XLAddress xlCellAddressLeft = new XLAddress(3, 3);
            XLAddress xlCellAddressRight = new XLAddress(4, 1);
            bool expected = true;
            bool actual;
            actual = (xlCellAddressLeft >= xlCellAddressRight);
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void op_GreaterThanOrEqualTest2()
        {
            XLAddress xlCellAddressLeft = new XLAddress(3, 3);
            XLAddress xlCellAddressRight = new XLAddress(3, 3);
            bool expected = true;
            bool actual;
            actual = (xlCellAddressLeft >= xlCellAddressRight);
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void op_GreaterThanOrEqualTest3()
        {
            XLAddress xlCellAddressLeft = new XLAddress(3, 3);
            XLAddress xlCellAddressRight = new XLAddress(3, 4);
            bool expected = false;
            bool actual;
            actual = (xlCellAddressLeft >= xlCellAddressRight);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        ///A test for op_Inequality
        ///</summary>
        [TestMethod()]
        public void op_InequalityTest()
        {
            XLAddress xlCellAddressLeft = new XLAddress(1,1); 
            XLAddress xlCellAddressRight = new XLAddress(1,2);
            bool expected = true;
            bool actual;
            actual = (xlCellAddressLeft != xlCellAddressRight);
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void op_InequalityTest1()
        {
            XLAddress xlCellAddressLeft = new XLAddress(1, 1);
            XLAddress xlCellAddressRight = new XLAddress(1, 1);
            bool expected = false;
            bool actual;
            actual = (xlCellAddressLeft != xlCellAddressRight);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        ///A test for op_LessThan
        ///</summary>
        [TestMethod()]
        public void op_LessThanTest()
        {
            XLAddress xlCellAddressLeft = new XLAddress(3, 3);
            XLAddress xlCellAddressRight = new XLAddress(2, 4);
            bool expected = true;
            bool actual;
            actual = (xlCellAddressLeft < xlCellAddressRight);
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void op_LessThanTest1()
        {
            XLAddress xlCellAddressLeft = new XLAddress(3, 3);
            XLAddress xlCellAddressRight = new XLAddress(4, 1);
            bool expected = true;
            bool actual;
            actual = (xlCellAddressLeft < xlCellAddressRight);
            Assert.AreEqual(expected, actual);
        }

        public void op_LessThanTest2()
        {
            XLAddress xlCellAddressLeft = new XLAddress(3, 3);
            XLAddress xlCellAddressRight = new XLAddress(3, 4);
            bool expected = true;
            bool actual;
            actual = (xlCellAddressLeft < xlCellAddressRight);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        ///A test for op_LessThanOrEqual
        ///</summary>
        [TestMethod()]
        public void op_LessThanOrEqualTest()
        {
            XLAddress xlCellAddressLeft = new XLAddress(3, 3);
            XLAddress xlCellAddressRight = new XLAddress(2, 4);
            bool expected = true;
            bool actual;
            actual = (xlCellAddressLeft <= xlCellAddressRight);
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void op_LessThanOrEqualTest1()
        {
            XLAddress xlCellAddressLeft = new XLAddress(3, 3);
            XLAddress xlCellAddressRight = new XLAddress(4, 1);
            bool expected = true;
            bool actual;
            actual = (xlCellAddressLeft <= xlCellAddressRight);
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void op_LessThanOrEqualTest2()
        {
            XLAddress xlCellAddressLeft = new XLAddress(3, 3);
            XLAddress xlCellAddressRight = new XLAddress(3, 3);
            bool expected = true;
            bool actual;
            actual = (xlCellAddressLeft <= xlCellAddressRight);
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void op_LessThanOrEqualTest3()
        {
            XLAddress xlCellAddressLeft = new XLAddress(3, 3);
            XLAddress xlCellAddressRight = new XLAddress(3, 4);
            bool expected = true;
            bool actual;
            actual = (xlCellAddressLeft <= xlCellAddressRight);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        ///A test for op_Subtraction
        ///</summary>
        [TestMethod()]
        public void op_SubtractionTest()
        {
            XLAddress xlCellAddressLeft = new XLAddress(6, 3);
            uint right = 2; 
            XLAddress expected = new XLAddress(4,1); 
            XLAddress actual;
            actual = (xlCellAddressLeft - right);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        ///A test for op_Subtraction
        ///</summary>
        [TestMethod()]
        public void op_SubtractionTest1()
        {
            XLAddress xlCellAddressLeft = new XLAddress(6,3);
            XLAddress xlCellAddressRight = new XLAddress(2,1);
            XLAddress expected = new XLAddress(4,2); 
            XLAddress actual;
            actual = (xlCellAddressLeft - xlCellAddressRight);
            Assert.AreEqual(expected, actual);
        }
    }
}
