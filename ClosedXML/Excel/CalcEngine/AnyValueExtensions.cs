using OneOf;
using System;
using System.Globalization;
using AnyValue = OneOf.OneOf<ClosedXML.Excel.CalcEngine.Logical, ClosedXML.Excel.CalcEngine.Number1, ClosedXML.Excel.CalcEngine.Text, ClosedXML.Excel.CalcEngine.Error1, ClosedXML.Excel.CalcEngine.Array, ClosedXML.Excel.CalcEngine.Reference1>;
using ScalarValue = OneOf.OneOf<ClosedXML.Excel.CalcEngine.Logical, ClosedXML.Excel.CalcEngine.Number1, ClosedXML.Excel.CalcEngine.Text, ClosedXML.Excel.CalcEngine.Error1>;
using AggregateValue = OneOf.OneOf<ClosedXML.Excel.CalcEngine.Array, ClosedXML.Excel.CalcEngine.Reference1>;
using System.Linq;

namespace ClosedXML.Excel.CalcEngine
{
    internal static class AnyValueExtensions
    {
        #region Type conversion functions
        public static bool TryPickScalar(this AnyValue value, out ScalarValue scalar, out AggregateValue aggregate)
        {
            scalar = value.Index switch
            {
                0 => value.AsT0,
                1 => value.AsT1,
                2 => value.AsT2,
                3 => value.AsT3,
                _ => default
            };
            aggregate = value.Index switch
            {
                4 => value.AsT4,
                5 => value.AsT5,
                _ => default
            };
            return value.Index <= 3;
        }

        public static AnyValue ToAnyValue(this ScalarValue scalar)
        {
            return scalar.Match(
                logical => AnyValue.FromT0(logical),
                number => AnyValue.FromT1(number),
                text => AnyValue.FromT2(text),
                error => AnyValue.FromT3(error));
        }

        #endregion

        #region Arithmetic unary operations

        public static AnyValue UnaryPlus(this AnyValue value)
        {
            // Unary plus doesn't even convert to number. Type and value is retained.
            return value;
        }

        public static AnyValue UnaryMinus(this AnyValue value, CalcContext context) => UnaryOperation(value, x => -x, context);

        public static AnyValue UnaryPercent(this AnyValue value, CalcContext context) => UnaryOperation(value, x => x / 100.0, context);

        public static AnyValue UnaryImplicitIntersection(this AnyValue value, CalcContext context) => throw new NotImplementedException();

        private static AnyValue UnaryOperation(AnyValue value, Func<Number1, Number1> f, CalcContext context)
        {
            if (value.TryPickScalar(out var scalar, out var aggregate))
            {
                return UnaryArithmeticOp(scalar, f, context.Converter).ToAnyValue();
            }

            return aggregate.Match(
                array => ApplyOnArray(array, arrayConst => UnaryArithmeticOp(arrayConst, f, context.Converter)),
                reference => ApplyOnReference(reference, cellValue => UnaryArithmeticOp(cellValue, f, context.Converter)));
        }

        private static AnyValue ApplyOnArray(Array array, Func<ScalarValue, ScalarValue> op)
        {
            var data = new ScalarValue[array.Height, array.Width];
            for (int y = 0; y < array.Height; ++y)
                for (int x = 0; x < array.Width; ++x)
                    data[y, x] = op(array[y, x]);
            return AnyValue.FromT4(new ConstArray(data));
        }

        private static AnyValue ApplyOnReference(Reference1 reference, Func<ScalarValue, ScalarValue> op)
        {
            if (reference.Areas.Count != 1)
                return Error1.CellValue;

            var area = reference.Areas.Single();
            var width = area.Width;
            var height = area.Height;
            var data = new ScalarValue[width, height];
            for (int y = 0; y < height; ++y)
            {
                for (int x = 0; x < width; ++x)
                {
                    var row = area.Row + y;
                    var column = area.Column + x;
                    var cellValue = GetCellValue(area, row, column);
                    data[y, x] = op(cellValue);
                }
            }
            return AnyValue.FromT4(new ConstArray(data));
        }

        private static ScalarValue UnaryArithmeticOp(ScalarValue value, Func<Number1, Number1> op, ValueConverter converter)
        {
            var conversionResult = value.Match(
                logical => converter.ToNumber(logical),
                number => number,
                text => converter.ToNumber(text),
                error => error);

            if (!conversionResult.TryPickT0(out var number, out var error))
                return error;

            return op(number);
        }

        #endregion

        #region Arithmetic binary operators

        public static AnyValue BinaryPlus(this AnyValue left, AnyValue right, CalcContext context)
        {
            BinaryNumberFunc f = (lhs, rhs) => new Number1(lhs + rhs);
            BinaryFunc g = (leftItem, rightItem) => BinaryArithmeticOp(leftItem, rightItem, f, context.Converter);
            return BinaryOperation(left, right, g, context);
        }

        public static AnyValue BinaryMinus(this AnyValue left, AnyValue right, CalcContext context)
        {
            BinaryNumberFunc f = (lhs, rhs) => new Number1(lhs - rhs);
            BinaryFunc g = (leftItem, rightItem) => BinaryArithmeticOp(leftItem, rightItem, f, context.Converter);
            return BinaryOperation(left, right, g, context);
        }

        public static AnyValue BinaryMult(this AnyValue left, AnyValue right, CalcContext context)
        {
            BinaryNumberFunc f = (lhs, rhs) => new Number1(lhs * rhs);
            BinaryFunc g = (leftItem, rightItem) => BinaryArithmeticOp(leftItem, rightItem, f, context.Converter);
            return BinaryOperation(left, right, g, context);
        }

        public static AnyValue BinaryDiv(this AnyValue left, AnyValue right, CalcContext context)
        {
            BinaryNumberFunc f = (lhs, rhs) => rhs == 0.0 ? Error1.DivZero : new Number1(lhs / rhs);
            BinaryFunc g = (leftItem, rightItem) => BinaryArithmeticOp(leftItem, rightItem, f, context.Converter);
            return BinaryOperation(left, right, g, context);
        }

        public static AnyValue BinaryExp(this AnyValue left, AnyValue right, CalcContext context)
        {
            BinaryNumberFunc f = (lhs, rhs) => lhs == 0 && rhs == 0 ? Error1.CellValue : new Number1(Math.Pow(lhs, rhs));
            BinaryFunc g = (leftItem, rightItem) => BinaryArithmeticOp(leftItem, rightItem, f, context.Converter);
            return BinaryOperation(left, right, g, context);
        }

        #endregion

        #region Comparison operators

        public static AnyValue IsEqual(this AnyValue left, AnyValue right, CalcContext context)
        {
            BinaryFunc g = (leftItem, rightItem) => CompareValues(leftItem, rightItem, context.Culture).Match<ScalarValue>(
                cmp => cmp == 0 ? Logical.True : Logical.False,
                error => error);
            return BinaryOperation(left, right, g, context);
        }

        public static AnyValue IsNotEqual(this AnyValue left, AnyValue right, CalcContext context)
        {
            BinaryFunc g = (leftItem, rightItem) => CompareValues(leftItem, rightItem, context.Culture).Match<ScalarValue>(
                cmp => cmp != 0 ? Logical.True : Logical.False,
                error => error);
            return BinaryOperation(left, right, g, context);
        }

        public static AnyValue IsGreaterThan(this AnyValue left, AnyValue right, CalcContext context)
        {
            BinaryFunc g = (leftItem, rightItem) => CompareValues(leftItem, rightItem, context.Culture).Match<ScalarValue>(
                cmp => cmp > 0 ? Logical.True : Logical.False,
                error => error);
            return BinaryOperation(left, right, g, context);
        }

        public static AnyValue IsGreaterThanOrEqual(this AnyValue left, AnyValue right, CalcContext context)
        {
            BinaryFunc g = (leftItem, rightItem) => CompareValues(leftItem, rightItem, context.Culture).Match<ScalarValue>(
                cmp => cmp >= 0 ? Logical.True : Logical.False,
                error => error);
            return BinaryOperation(left, right, g, context);
        }

        public static AnyValue IsLessThan(this AnyValue left, AnyValue right, CalcContext context)
        {
            BinaryFunc g = (leftItem, rightItem) => CompareValues(leftItem, rightItem, context.Culture).Match<ScalarValue>(
                cmp => cmp < 0 ? Logical.True : Logical.False,
                error => error);
            return BinaryOperation(left, right, g, context);
        }

        public static AnyValue IsLessThanOrEqual(this AnyValue left, AnyValue right, CalcContext context)
        {
            BinaryFunc g = (leftItem, rightItem) => CompareValues(leftItem, rightItem, context.Culture).Match<ScalarValue>(
                cmp => cmp <= 0 ? Logical.True : Logical.False,
                error => error);
            return BinaryOperation(left, right, g, context);
        }

        #endregion

        private static AnyValue BinaryOperation(AnyValue left, AnyValue right, BinaryFunc func, CalcContext context)
        {
            var isLeftScalar = left.TryPickScalar(out var leftScalar, out var leftAggregate);
            var isRightScalar = right.TryPickScalar(out var rightScalar, out var rightAggregate);

            if (isLeftScalar && isRightScalar)
                return func(leftScalar, rightScalar).ToAnyValue();

            if (isLeftScalar)
            {
                return rightAggregate.Match(
                    array => ApplyOnArray(
                        new ScalarArray(leftScalar, array.Width, array.Height),
                        array,
                        func),
                    reference => throw new NotImplementedException());
            }

            if (isRightScalar)
            {
                return leftAggregate.Match(
                    array => ApplyOnArray(
                        array,
                        new ScalarArray(rightScalar, array.Width, array.Height),
                        func),
                    reference => throw new NotImplementedException());
            }

            // Both are aggregates
            return leftAggregate.Match(
                leftArray => rightAggregate.Match(
                        rightArray =>
                        {
                            var width = Math.Max(leftArray.Width, rightArray.Width);
                            var height = Math.Max(leftArray.Height, rightArray.Height);
                            return ApplyOnArray(
                                new ResizedArray(leftArray, width, height),
                                new ResizedArray(rightArray, width, height),
                                func);
                        },
                        rightReference => throw new NotImplementedException()),
                leftReference => throw new NotImplementedException());
        }

        private static AnyValue ApplyOnArray(Array leftArray, Array rightArray, BinaryFunc func)
        {
            if (leftArray.Width != rightArray.Width || leftArray.Height != rightArray.Height)
                throw new ArgumentException("Array dimensions differ.");

            var data = new ScalarValue[leftArray.Height, leftArray.Width];
            for (int y = 0; y < leftArray.Height; ++y)
                for (int x = 0; x < leftArray.Width; ++x)
                {
                    var leftItem = leftArray[y, x];
                    var rightItem = rightArray[y, x];
                    data[y, x] = func(leftItem, rightItem);
                }
            return AnyValue.FromT4(new ConstArray(data));
        }

        private static ScalarValue BinaryArithmeticOp(ScalarValue lhs, ScalarValue rhs, BinaryNumberFunc func, ValueConverter converter)
        {
            var leftConversionResult = lhs.CovertToNumber(converter);
            if (!leftConversionResult.TryPickT0(out var leftNumber, out var leftError))
            {
                return leftError;
            }

            var rightConversionResult = rhs.CovertToNumber(converter);
            if (!rightConversionResult.TryPickT0(out var rightNumber, out var rightError))
            {
                return rightError;
            }

            return func(leftNumber, rightNumber).Match(
                number => ScalarValue.FromT1(number),
                error => ScalarValue.FromT3(error));
        }

        private static OneOf<Number1, Error1> CovertToNumber(this ScalarValue value, ValueConverter converter)
        {
            return value.Match(
                logical => converter.ToNumber(logical),
                number => number,
                text => converter.ToNumber(text),
                error => error);
        }

        /// <summary>
        /// Compare two scalar values using Excel semantic. Rules for comparison are following:
        /// <list type="bullet">
        ///     <item>Logical is always greater than any text (thus transitively greater than any number)</item>
        ///     <item>Text is always greater than any number, even if empty string</item>
        ///     <item>Logical are compared by value</item>
        ///     <item>Numbers are compared by value</item>
        ///     <item>Text is compared by through case insensitive comparison for workbook culture.</item>
        ///     <item>
        ///         If any argument is error, return error (general rule for all operators).
        ///         If all args are errors, pick leftmost error (technically it is left to
        ///         implementation, but excel sems to be using left one)
        ///     </item>
        /// </list>
        /// </summary>
        /// <param name="lhs">Left hand operand of the comparison.</param>
        /// <param name="rhs">Right hand operand of the comparison.</param>
        /// <param name="culture">Culture to use for comparison.</param>
        /// <returns>
        ///     Return -1 (negative)  if left less than right
        ///     Return 1 (positive) if right greater than left
        ///     Return 0 if both operands are considered equal.
        /// </returns>
        private static OneOf<int, Error1> CompareValues(ScalarValue lhs, ScalarValue rhs, CultureInfo culture)
        {
            return lhs.Match(
                leftLogical => rhs.Match<OneOf<int, Error1>>(
                        rightLogical => leftLogical.Value.CompareTo(rightLogical.Value),
                        rightNumber => -1,
                        rightText => -1,
                        rightError => rightError),
                leftNumber => rhs.Match<OneOf<int, Error1>>(
                        rightLogical => 1,
                        rightNumber => leftNumber.Value.CompareTo(rightNumber.Value),
                        rightText => 1,
                        rightError => rightError),
                leftText => rhs.Match<OneOf<int, Error1>>(
                        rightLogical => 1,
                        rightNumber => -1,
                        rightText => string.Compare(leftText.Value, rightText.Value, culture, CompareOptions.IgnoreCase),
                        rightError => rightError),
                leftError => leftError);
        }

        private static ScalarValue GetCellValue(Area area, int row, int column)
        {
            var value = area.Worksheet.GetCellValue(row, column);
            if (value is bool boolValue)
                return ScalarValue.FromT0(new Logical(boolValue));
            if (value is double numberValue)
                return ScalarValue.FromT1(new Number1(numberValue));
            if (value is string stringValue)
            {
                return stringValue == string.Empty
                    ? ScalarValue.FromT1(new Number1(0))
                    : ScalarValue.FromT2(new Text(stringValue));
            }

            throw new NotImplementedException("Not sure how to get error from a cell.");
        }

        private delegate ScalarValue BinaryFunc(ScalarValue lhs, ScalarValue rhs);

        private delegate OneOf<Number1, Error1> BinaryNumberFunc(Number1 lhs, Number1 rhs);
    }
}
