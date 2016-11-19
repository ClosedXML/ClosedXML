using System;
using System.Runtime.InteropServices;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML_Tests.Excel.CalcEngine
{

    [TestFixture]
    public class InformationTests
    {

        #region IsBlank Tests
        [Test]
        public void IsBlank_Single_true()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet");
            var actual = ws.Evaluate("=IsBlank(A1)");
            Assert.AreEqual(true, actual);
        }

        [Test]
        public void IsBlank_Single_false()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet");
            ws.Cell("A1").Value = " ";
            var actual = ws.Evaluate("=IsBlank(A1)");
            Assert.AreEqual(false, actual);
        }

        [Test]
        public void IsBlank_MultipleAllEmpty_true()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet");
            var actual = ws.Evaluate("=IsBlank(A1:A3)");
            Assert.AreEqual(true,actual);

        }

        [Test]
        public void IsBlank_MultipleAllFill_false()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet");
            ws.Cell("A1").Value = "1";
            ws.Cell("A2").Value = "1";
            ws.Cell("A3").Value = "1";
            var actual = ws.Evaluate("=IsBlank(A1:A3)");
            Assert.AreEqual(false,actual);
        }

        [Test]
        public void IsBlank_MultipleMixedFill_false()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet");
            ws.Cell("A1").Value = "1";
            ws.Cell("A3").Value = "1";
            var actual = ws.Evaluate("=IsBlank(A1:A3)");
            Assert.AreEqual(false, actual);
        }
        #endregion

        #region IsEven Tests

        [Test]
        public void IsEven_Single_False()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet");

            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = 1.2;
            ws.Cell("A3").Value = 3;

            var actual = ws.Evaluate("=IsEven(A1)");
            Assert.AreEqual(false,actual);

            actual = ws.Evaluate("=IsEven(A2)");
            Assert.AreEqual(false, actual);

            actual = ws.Evaluate("=IsEven(A3)");
            Assert.AreEqual(false, actual);
        }

        [Test]
        public void IsEven_Single_True()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet");

            ws.Cell("A1").Value = 4;
            ws.Cell("A2").Value = 0.2;
            ws.Cell("A3").Value = 12.2;

            var actual = ws.Evaluate("=IsEven(A1)");
            Assert.AreEqual(true, actual);

            actual = ws.Evaluate("=IsEven(A2)");
            Assert.AreEqual(true, actual);

            actual = ws.Evaluate("=IsEven(A3)");
            Assert.AreEqual(true, actual);
        }
        #endregion

        #region IsLogical Tests

        [Test]
        public void IsLogical_Simple_True()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet");

            ws.Cell("A1").Value = true;

            var actual = ws.Evaluate("=IsLogical(A1)");
            Assert.AreEqual(true, actual);
        }

        [Test]
        public void IsLogical_Simpe_False()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet");

            ws.Cell("A1").Value = 123;

            var actual = ws.Evaluate("=IsLogical(A1)");
            Assert.AreEqual(false, actual);
        }
        #endregion

        #region IsNotText Tests

        [Test]
        public void IsNotText_Simple_true()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet");
            ws.Cell("A1").Value = "123"; //Double Value
            ws.Cell("A2").Value = DateTime.Now; //Date Value
            ws.Cell("A3").Value = "12,235.5"; //Coma Formatting
            ws.Cell("A4").Value = "$12,235.5"; //Currency Value
            ws.Cell("A5").Value = true; //Bool Value
            ws.Cell("A6").Value = "12%"; //Percentage Value

            var actual = ws.Evaluate("=IsNonText(A1)");
            Assert.AreEqual(true, actual);
            actual = ws.Evaluate("=IsNonText(A2)");
            Assert.AreEqual(true, actual);
            actual = ws.Evaluate("=IsNonText(A3)");
            Assert.AreEqual(true, actual);
            actual = ws.Evaluate("=IsNonText(A4)");
            Assert.AreEqual(true, actual);
            actual = ws.Evaluate("=IsNonText(A5)");
            Assert.AreEqual(true, actual);
            actual = ws.Evaluate("=IsNonText(A6)");
            Assert.AreEqual(true, actual);
        }

        [Test]
        public void IsNotText_Simple_false()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet");
            ws.Cell("A1").Value = "asd";
            var actual = ws.Evaluate("=IsNonText(A1)");
            Assert.AreEqual(false, actual);
        }

        #endregion

        #region IsNumber Tests

        [Test]
        public void IsNumber_Simple_true()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet");
            ws.Cell("A1").Value = "123"; //Double Value
            ws.Cell("A2").Value = DateTime.Now; //Date Value
            ws.Cell("A3").Value = "12,235.5"; //Coma Formatting
            ws.Cell("A4").Value = "$12,235.5"; //Currency Value
            ws.Cell("A5").Value = "12%"; //Percentage Value

            var actual = ws.Evaluate("=IsNumber(A1)");
            Assert.AreEqual(true, actual);
            actual = ws.Evaluate("=IsNumber(A2)");
            Assert.AreEqual(true, actual);
            actual = ws.Evaluate("=IsNumber(A3)");
            Assert.AreEqual(true, actual);
            actual = ws.Evaluate("=IsNumber(A4)");
            Assert.AreEqual(true, actual);
            actual = ws.Evaluate("=IsNumber(A5)");
            Assert.AreEqual(true, actual);
        }

        [Test]
        public void IsNumber_Simple_false()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet");
            ws.Cell("A1").Value = "asd"; //String Value
            ws.Cell("A2").Value = true; //Bool Value

            var actual = ws.Evaluate("=IsNumber(A1)");
            Assert.AreEqual(false, actual);
            actual = ws.Evaluate("=IsNumber(A2)");
            Assert.AreEqual(false, actual);
        }
        #endregion

        #region IsOdd Test

        [Test]
        public void IsOdd_Simple_true()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet");

            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = 1.2;
            ws.Cell("A3").Value = 3;

            var actual = ws.Evaluate("=IsOdd(A1)");
            Assert.AreEqual(true, actual);
            actual = ws.Evaluate("=IsOdd(A2)");
            Assert.AreEqual(true, actual);
            actual = ws.Evaluate("=IsOdd(A3)");
            Assert.AreEqual(true, actual);
        }

        [Test]
        public void IsOdd_Simple_false()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet");

            ws.Cell("A1").Value = 4;
            ws.Cell("A2").Value = 0.2;
            ws.Cell("A3").Value = 12.2;

            var actual = ws.Evaluate("=IsOdd(A1)");
            Assert.AreEqual(false, actual);
            actual = ws.Evaluate("=IsOdd(A2)");
            Assert.AreEqual(false, actual);
            actual = ws.Evaluate("=IsOdd(A3)");
            Assert.AreEqual(false, actual);
        }
        #endregion

        #region IsText Tests
        [Test]
        public void IsText_Simple_true()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet");

            ws.Cell("A1").Value = "asd";

            var actual = ws.Evaluate("=IsText(A1)");
            Assert.AreEqual(true, actual);
        }

        [Test]
        public void IsText_Simple_false()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet");
            ws.Cell("A1").Value = "123"; //Double Value
            ws.Cell("A2").Value = DateTime.Now; //Date Value
            ws.Cell("A3").Value = "12,235.5"; //Coma Formatting
            ws.Cell("A4").Value = "$12,235.5"; //Currency Value
            ws.Cell("A5").Value = true; //Bool Value
            ws.Cell("A6").Value = "12%"; //Percentage Value

            var actual = ws.Evaluate("=IsText(A1)");
            Assert.AreEqual(false, actual);
            actual = ws.Evaluate("=IsText(A2)");
            Assert.AreEqual(false, actual);
            actual = ws.Evaluate("=IsText(A3)");
            Assert.AreEqual(false, actual);
            actual = ws.Evaluate("=IsText(A4)");
            Assert.AreEqual(false, actual);
            actual = ws.Evaluate("=IsText(A5)");
            Assert.AreEqual(false, actual);
            actual = ws.Evaluate("=IsText(A6)");
            Assert.AreEqual(false, actual);
        }
        #endregion

        #region N Tests

        [Test]
        public void N_Date_SerialNumber()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet");
            var testedDate = DateTime.Now;
            ws.Cell("A1").Value = testedDate;
            var actual = ws.Evaluate("=N(A1)");
            Assert.AreEqual(testedDate.ToOADate(),actual);
        }

        [Test]
        public void N_Number_Number()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet");
            var testedValue = 123;
            ws.Cell("A1").Value = testedValue;
            var actual = ws.Evaluate("=N(A1)");
            Assert.AreEqual(testedValue, actual);
        }

        [Test]
        public void N_String_Zero()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet");
            ws.Cell("A1").Value = "asd";
            var actual = ws.Evaluate("=N(A1)");
            Assert.AreEqual(0, actual);
        }

        [Test]
        public void N_True_One()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet");
            ws.Cell("A1").Value = true;
            var actual = ws.Evaluate("=N(A1)");
            Assert.AreEqual(1, actual);
        }

        [Test]
        public void N_False_Zero()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet");
            ws.Cell("A1").Value = false;
            var actual = ws.Evaluate("=N(A1)");
            Assert.AreEqual(0, actual);
        }
        #endregion


    }
}