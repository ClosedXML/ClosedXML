using System;
using System.Linq;
using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClosedXML_Tests
{
    ///<summary>
    ///	This is a test class for XLRichStringTest and is intended
    ///	to contain all XLRichStringTest Unit Tests
    ///</summary>
    [TestClass]
    public class XLRichStringTest
    {
        private TestContext testContextInstance;

        ///<summary>
        ///	Gets or sets the test context which provides
        ///	information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get { return testContextInstance; }
            set { testContextInstance = value; }
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
        ///<summary>
        ///	A test for ToString
        ///</summary>
        [TestMethod]
        public void ToStringTest()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).RichText;

            richString.AddText("Hello");
            richString.AddText(" ");
            richString.AddText("World");
            string expected = "Hello World";
            string actual = richString.ToString();
            Assert.AreEqual(expected, actual);

            richString.AddText("!");
            expected = "Hello World!";
            actual = richString.ToString();
            Assert.AreEqual(expected, actual);

            richString.Clear();
            expected = String.Empty;
            actual = richString.ToString();
            Assert.AreEqual(expected, actual);
        }

        ///<summary>
        ///	A test for AddText
        ///</summary>
        [TestMethod]
        public void AddTextTest1()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            var richString = cell.RichText;

            string text = "Hello";
            richString.AddText(text).SetBold().SetFontColor(XLColor.Red);

            Assert.AreEqual(cell.GetString(), text);
            Assert.AreEqual(cell.RichText.First().Bold, true);
            Assert.AreEqual(cell.RichText.First().FontColor, XLColor.Red);

            Assert.AreEqual(1, richString.Count);

            richString.AddText("World");
            Assert.AreEqual(richString.First().Text, text, "Item in collection is not the same as the one returned");
        }

        [TestMethod]
        public void AddTextTest2()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            Int32 number = 123;

            cell.SetValue(number).Style
                    .Font.SetBold()
                    .Font.SetFontColor(XLColor.Red);

            string text = number.ToString();

            Assert.AreEqual(cell.RichText.ToString(), text);
            Assert.AreEqual(cell.RichText.First().Bold, true);
            Assert.AreEqual(cell.RichText.First().FontColor, XLColor.Red);

            Assert.AreEqual(1, cell.RichText.Count);

            cell.RichText.AddText("World");
            Assert.AreEqual(cell.RichText.First().Text, text, "Item in collection is not the same as the one returned");
        }

        [TestMethod]
        public void AddTextTest3()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            Int32 number = 123;
            cell.Value = number;
            cell.Style
                    .Font.SetBold()
                    .Font.SetFontColor(XLColor.Red);

            string text = number.ToString();

            Assert.AreEqual(cell.RichText.ToString(), text);
            Assert.AreEqual(cell.RichText.First().Bold, true);
            Assert.AreEqual(cell.RichText.First().FontColor, XLColor.Red);

            Assert.AreEqual(1, cell.RichText.Count);

            cell.RichText.AddText("World");
            Assert.AreEqual(cell.RichText.First().Text, text, "Item in collection is not the same as the one returned");
        }

        [TestMethod]
        public void HasRichTextTest1()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.RichText.AddText("123");

            Assert.AreEqual(true, cell.HasRichText);

            cell.DataType = XLCellValues.Text;

            Assert.AreEqual(true, cell.HasRichText);

            cell.DataType = XLCellValues.Number;

            Assert.AreEqual(false, cell.HasRichText);

            cell.RichText.AddText("123");

            Assert.AreEqual(true, cell.HasRichText);

            cell.Value = 123;

            Assert.AreEqual(false, cell.HasRichText);

            cell.RichText.AddText("123");

            Assert.AreEqual(true, cell.HasRichText);

            cell.SetValue("123");

            Assert.AreEqual(false, cell.HasRichText);
        }

        [TestMethod]
        public void AccessRichTextTest1()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.RichText.AddText("12");
            cell.DataType = XLCellValues.Number;

            Assert.AreEqual(12.0, cell.GetDouble());

            var richText = cell.RichText;

            Assert.AreEqual("12", richText.ToString());

            richText.AddText("34");

            Assert.AreEqual("1234", cell.GetString());

            Assert.AreEqual(XLCellValues.Text, cell.DataType);

            cell.DataType = XLCellValues.Number;

            Assert.AreEqual(1234.0, cell.GetDouble());
        }

        ///<summary>
        ///	A test for Characters
        ///</summary>
        [TestMethod]
        public void CharactersTest()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).RichText;

            int index = 0; // TODO: Initialize to an appropriate value
            int length = 0; // TODO: Initialize to an appropriate value
            IXLRichText expected = null; // TODO: Initialize to an appropriate value
            IXLRichText actual;
            actual = richString.Characters(index, length);
            Assert.AreEqual(expected, actual);
            Assert.Inconclusive("Verify the correctness of this test method.");
        }

        ///<summary>
        ///	A test for Clear
        ///</summary>
        [TestMethod]
        public void ClearTest()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).RichText;

            richString.AddText("Hello");
            richString.AddText(" ");
            richString.AddText("World!");

            richString.Clear();
            String expected = String.Empty;
            String actual = richString.ToString();
            Assert.AreEqual(expected, actual);

            Assert.AreEqual(0, richString.Count);
        }

        [TestMethod]
        public void CountTest()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).RichText;

            richString.AddText("Hello");
            richString.AddText(" ");
            richString.AddText("World!");

            Assert.AreEqual(3, richString.Count);
        }
    }
}