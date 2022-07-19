using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine;
using NUnit.Framework;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using AnyValue = OneOf.OneOf<ClosedXML.Excel.CalcEngine.Logical, ClosedXML.Excel.CalcEngine.Number1, ClosedXML.Excel.CalcEngine.Text, ClosedXML.Excel.CalcEngine.Error1, ClosedXML.Excel.CalcEngine.Array, ClosedXML.Excel.CalcEngine.Reference>;
using ScalarValue = OneOf.OneOf<ClosedXML.Excel.CalcEngine.Logical, ClosedXML.Excel.CalcEngine.Number1, ClosedXML.Excel.CalcEngine.Text, ClosedXML.Excel.CalcEngine.Error1>;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    public class OperationsTests
    {
        [Test]
        public void UnaryPlus_DoesntChangeTypeOrValue()
        {
        }

        [Test]
        public void UnaryMinus_ConvertsScalarToNumberAndChangesSign()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var calcContext = new CalcContext(System.Globalization.CultureInfo.InvariantCulture, (XLWorksheet)ws);

            // Logical is converted 
            AnyValue logical = new Logical(true);
            var resultLogical = logical.UnaryMinus(calcContext);
            Assert.AreEqual(AnyValue.FromT1(new Number1(-1)), resultLogical);

            AnyValue number = new Text("1.5");
            var resultNumber = number.UnaryMinus(calcContext);
            Assert.AreEqual(AnyValue.FromT1(new Number1(-1.5)), resultNumber);

            AnyValue text = new Text("-1");
            var resultText = text.UnaryMinus(calcContext);
            Assert.AreEqual(AnyValue.FromT1(new Number1(1)), resultText);

            AnyValue error = Error1.DivZero;
            var resultError = error.UnaryMinus(calcContext);
            Assert.AreEqual(AnyValue.FromT3(Error1.DivZero), resultError);

            var a = new ScalarValue[,] { { ScalarValue.FromT1(new Number1(1)), ScalarValue.FromT3(Error1.DivZero) } };
            AnyValue array = new ConstArray(a);
            var resultArray = array.UnaryMinus(calcContext);
            var b = new ScalarValue[,] { { ScalarValue.FromT1(new Number1(-1)), ScalarValue.FromT3(Error1.DivZero) } };
            var c = (ConstArray)resultArray.AsT4;
            Assert.AreEqual(b, c._data);
        }


        [Test]
        public void BinaryPlus_ConvertsScalarToNumberAndChangesSign()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var calcContext = new CalcContext(System.Globalization.CultureInfo.InvariantCulture, (XLWorksheet)ws);

            // Logical is converted 
            AnyValue logical = new Logical(true);
            AnyValue number = new Number1(1.5);
            AnyValue numberText = new Text("1.75");
            AnyValue nonNumberText = new Text("abc");
            AnyValue error = Error1.DivZero;

            var resultOfSameType = number.BinaryPlus(number, calcContext);
            Assert.AreEqual(AnyValue.FromT1(new Number1(3)), resultOfSameType);

            var resultWithBothConversions = numberText.BinaryMinus(logical, calcContext);
            Assert.AreEqual(AnyValue.FromT1(new Number1(0.75)), resultWithBothConversions);

            var oneConversionFailsResult = logical.BinaryMinus(nonNumberText, calcContext);
            Assert.AreEqual(AnyValue.FromT3(Error1.Value), oneConversionFailsResult);

            var array2x1 = new ConstArray(new ScalarValue[,] { { ScalarValue.FromT1(new Number1(1)), ScalarValue.FromT3(Error1.DivZero) } });

            var arrayWithConstResult = AnyValue.FromT4(array2x1).BinaryPlus(number, calcContext);

            var array1x2 = new ConstArray(new ScalarValue[,] { { ScalarValue.FromT1(new Number1(1)) }, { ScalarValue.FromT2(new Text("14")) } });
            var arrayWithArrayResult = AnyValue.FromT4(array2x1).BinaryPlus(array1x2, calcContext);

            //Assert.AreEqual(AnyValue.FromT3(Error1.CellValue), oneConversionFailsResult);

            /*
            var resultText = numberText.UnaryMinus(calcContext);
            Assert.AreEqual(AnyValue.FromT1(new Number1(1)), resultText);

            var resultError = error.UnaryMinus(calcContext);
            Assert.AreEqual(AnyValue.FromT3(Error1.DivZero), resultError);

            var a = new ScalarValue[,] { { ScalarValue.FromT1(new Number1(1)), ScalarValue.FromT3(Error1.DivZero) } };
            AnyValue array = new ConstArray(a);
            var resultArray = array.UnaryMinus(calcContext);
            var b = new ScalarValue[,] { { ScalarValue.FromT1(new Number1(-1)), ScalarValue.FromT3(Error1.DivZero) } };
            var c = (ConstArray)resultArray.AsT4;
            Assert.AreEqual(b, c._data);*/
        }

        [Test]
        public void callFuncTest()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var calcContext = new CalcContext(System.Globalization.CultureInfo.InvariantCulture, (XLWorksheet)ws);

            var dict = new FunctionRegistry();
            TestFuncRegistry.Register(dict);
            dict.TryGetFunc("TYPE", out FormulaFunction f);
            var result = f.CallFunction(calcContext, new Number1(5));
            Assert.AreEqual(AnyValue.FromT1(new Number1(2)), result);
        }


        [Test]
        public void CalcFunctionPrimitive()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var calcContext = new CalcContext(System.Globalization.CultureInfo.InvariantCulture, (XLWorksheet)ws);

            var dict = new FunctionRegistry();
            TestFuncRegistry.Register(dict);
            dict.TryGetFunc("SIN", out FormulaFunction f);
            var result = f.CallFunction(calcContext, new Number1(System.Math.PI / 2));
            Assert.AreEqual(AnyValue.FromT1(new Number1(1)), result);
        }

        internal static class TestFuncRegistry
        {
            public static void Register(FunctionRegistry registry)
            {
                registry.RegisterFunction("TYPE", typeof(TestFuncRegistry).GetMethod(nameof(Type1), BindingFlags.NonPublic | BindingFlags.Static));
                registry.RegisterFunction("SIN", typeof(TestFuncRegistry).GetMethod(nameof(Sin), BindingFlags.NonPublic | BindingFlags.Static));
                registry.RegisterFunction("COS", typeof(TestFuncRegistry).GetMethod(nameof(Cos), BindingFlags.NonPublic | BindingFlags.Static));
            }

            private static AnyValue Type1(CalcContext ctx, AnyValue value)
            {
                return value.Match<Number1>(
                    logical => 1,
                    number => 2,
                    text => 4,
                    error => 16,
                    array => 64,
                    reference => 64);
            }

            private static AnyValue Sin(CalcContext ctx, Number1 value)
            {
                return new Number1(System.Math.Sin(value));
            }

            private static AnyValue Cos(CalcContext ctx, Number1 value)
            {
                return new Number1(System.Math.Cos(value));
            }
        }

    }
}
