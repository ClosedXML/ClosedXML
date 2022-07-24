using OneOf;
using System;
using System.Globalization;
using AnyValue = OneOf.OneOf<ClosedXML.Excel.CalcEngine.Logical, ClosedXML.Excel.CalcEngine.Number1, ClosedXML.Excel.CalcEngine.Text, ClosedXML.Excel.CalcEngine.Error1, ClosedXML.Excel.CalcEngine.Array, ClosedXML.Excel.CalcEngine.Reference>;
using ScalarValue = OneOf.OneOf<ClosedXML.Excel.CalcEngine.Logical, ClosedXML.Excel.CalcEngine.Number1, ClosedXML.Excel.CalcEngine.Text, ClosedXML.Excel.CalcEngine.Error1>;
using AggregateValue = OneOf.OneOf<ClosedXML.Excel.CalcEngine.Array, ClosedXML.Excel.CalcEngine.Reference>;
using System.Linq;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// Reference operations.
    /// </summary>
    internal static class RefExt
    {
        public static AnyValue ImplicitIntersection(this AnyValue value, CalcContext context)
        {
            return value.Match(
                logical => logical,
                number => number,
                text => text,
                logical => logical,
                array => array,
                reference =>
                {
                    if (reference.IsSingleCell())
                    {
                        return reference;
                    }

                    return reference
                        .ImplicitIntersection(context.FormulaAddress)
                        .Match<AnyValue>(
                            singleCellReference => singleCellReference,
                            error => error);
                });
        }

        public static AnyValue ReferenceRange(this AnyValue left, AnyValue right)
        {
            return ReferenceOp(left, right, Reference.RangeOp);
        }

        public static AnyValue ReferenceUnion(this AnyValue left, AnyValue right)
        {
            return ReferenceOp(left, right, (leftRef, rightRef) => Reference.UnionOp(leftRef, rightRef));
        }

        private static AnyValue ReferenceOp(AnyValue left, AnyValue right, Func<Reference, Reference, OneOf<Reference, Error1>> fn)
        {
            var leftConversionResult = ConvertToReference(left);
            if (!leftConversionResult.TryPickT0(out var leftReference, out var leftError))
                return leftError;

            var rightConversionResult = ConvertToReference(right);
            if (!rightConversionResult.TryPickT0(out var rightReference, out var rightError))
                return rightError;

            return fn(leftReference, rightReference).Match<AnyValue>(
                reference => reference,
                error => error);
        }

        private static OneOf<Reference, Error1> ConvertToReference(AnyValue left)
        {
            return left.Match<OneOf<Reference, Error1>>(
                logical => Error1.Value,
                number => Error1.Value,
                text => Error1.Value,
                error => error,
                array => Error1.Value,
                reference => reference);
        }
    }

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

        public static AnyValue ToAnyValue(this AggregateValue aggregate)
        {
            return aggregate.Match(
                array => AnyValue.FromT4(array),
                reference => AnyValue.FromT5(reference));
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

        private static AnyValue UnaryOperation(AnyValue value, Func<Number1, Number1> f, CalcContext context)
        {
            if (value.TryPickScalar(out var scalar, out var aggregate))
            {
                return UnaryArithmeticOp(scalar, f, context.Converter).ToAnyValue();
            }

            return aggregate.Match(
                array => ApplyOnArray(array, arrayConst => UnaryArithmeticOp(arrayConst, f, context.Converter)),
                reference => ApplyOnReference(reference, cellValue => UnaryArithmeticOp(cellValue, f, context.Converter), context));
        }

        private static AnyValue ApplyOnArray(Array array, Func<ScalarValue, ScalarValue> op)
        {
            var data = new ScalarValue[array.Height, array.Width];
            for (int y = 0; y < array.Height; ++y)
                for (int x = 0; x < array.Width; ++x)
                    data[y, x] = op(array[y, x]);
            return AnyValue.FromT4(new ConstArray(data));
        }

        private static AnyValue ApplyOnReference(Reference reference, Func<ScalarValue, ScalarValue> op, CalcContext context)
        {
            if (reference.Areas.Count != 1)
                return Error1.Value;

            var area = reference.Areas.Single();
            var width = area.ColumnSpan;
            var height = area.RowSpan;
            var startColumn = area.FirstAddress.ColumnNumber;
            var startRow = area.FirstAddress.RowNumber;
            var data = new ScalarValue[width, height];
            for (int y = 0; y < height; ++y)
            {
                for (int x = 0; x < width; ++x)
                {
                    var row = startRow + y;
                    var column = startColumn + x;
                    var cellValue = GetCellValue(area, row, column, context);
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
            BinaryNumberFunc f = (lhs, rhs) => lhs == 0 && rhs == 0 ? Error1.Value : new Number1(Math.Pow(lhs, rhs));
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

        public static AnyValue Concat(this AnyValue left, AnyValue right, CalcContext context)
        {
            BinaryFunc g = (lhs, rhs) =>
            {
                return context.Converter.ToText(lhs).Match(
                    leftText => context.Converter.ToText(rhs).Match<OneOf<Text, Error1>>(rightText => new Text(leftText + rightText), rightError => rightError),
                    leftError => leftError).Match<ScalarValue>(text => text, error => error);
            };

            return BinaryOperation(left, right, g, context);
        }

        private static AnyValue BinaryOperation(AnyValue left, AnyValue right, BinaryFunc func, CalcContext context)
        {
            var isLeftScalar = left.TryPickScalar(out var leftScalar, out var leftAggregate);
            var isRightScalar = right.TryPickScalar(out var rightScalar, out var rightAggregate);

            if (isLeftScalar && isRightScalar)
                return func(leftScalar, rightScalar).ToAnyValue();

            // This is for dynamic arrays
            if (isLeftScalar)
            {
                return rightAggregate.Match(
                    array => ApplyOnArray(
                        new ScalarArray(leftScalar, array.Width, array.Height),
                        array,
                        func),
                    rightReference =>
                    {
                        if (rightReference.IsSingleCell())
                        {
                            var rightArea = rightReference.Areas.Single();
                            var rightCellValue = context.GetCellValue(rightArea.Worksheet, rightArea.FirstAddress.RowNumber, rightArea.FirstAddress.ColumnNumber);
                            return func(rightCellValue, rightScalar).ToAnyValue();
                        }

                        var referenceArrayResult = rightReference.ToArray(context);
                        if (!referenceArrayResult.TryPickT0(out var rightRefArray, out var rightError))
                            return rightError;

                        return ApplyOnArray(new ScalarArray(leftScalar, rightRefArray.Width, rightRefArray.Height), rightRefArray, func);
                    });
            }

            if (isRightScalar)
            {
                return leftAggregate.Match(
                    leftArray => ApplyOnArray(
                        leftArray,
                        new ScalarArray(rightScalar, leftArray.Width, leftArray.Height),
                        func),
                    leftReference =>
                    {
                        if (leftReference.IsSingleCell())
                        {
                            var leftArea = leftReference.Areas.Single();
                            var leftCellValue = context.GetCellValue(leftArea.Worksheet, leftArea.FirstAddress.RowNumber, leftArea.FirstAddress.ColumnNumber);
                            return func(leftCellValue, rightScalar).ToAnyValue();
                        }

                        var referenceArrayResult = leftReference.ToArray(context);
                        if (!referenceArrayResult.TryPickT0(out var leftRefArray, out var leftError))
                            return leftError;

                        return ApplyOnArray(leftRefArray, new ScalarArray(rightScalar, leftRefArray.Width, leftRefArray.Height), func);
                    });
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
                        rightReference =>
                        {
                            throw new NotImplementedException();
                        }),
                leftReference => rightAggregate.Match(
                        rightArray =>
                        {
                            throw new NotImplementedException();
                        },
                        rightReference =>
                        {
                            if (leftReference.Areas.Count > 1 || rightReference.Areas.Count > 1)
                                return Error1.Value;

                            var leftArea = leftReference.Areas.Single();
                            var rightArea = rightReference.Areas.Single();
                            var colSpan = Math.Max(leftArea.ColumnSpan, rightArea.ColumnSpan);
                            var rowSpan = Math.Max(leftArea.RowSpan, rightArea.RowSpan);
                            if (colSpan == 1 && rowSpan == 1)
                            {
                                var leftCellValue = context.GetCellValue(leftArea.Worksheet, leftArea.FirstAddress.RowNumber, leftArea.FirstAddress.ColumnNumber);
                                var rightCellValue = context.GetCellValue(rightArea.Worksheet, rightArea.FirstAddress.RowNumber, rightArea.FirstAddress.ColumnNumber);
                                return func(leftCellValue, rightCellValue).ToAnyValue();
                            }

                            return ApplyOnArray(
                                new ResizedArray(new ReferenceArray(leftArea, context), colSpan, rowSpan),
                                new ResizedArray(new ReferenceArray(rightArea, context), colSpan, rowSpan),
                                func);
                        }));
        }

        // If not a single area, error
        public static OneOf<Array, Error1> ToArray(this Reference reference, CalcContext context)
        {
            if (reference.Areas.Count != 1)
                throw new NotImplementedException();

            var area = reference.Areas.Single();

            return new ReferenceArray(area, context);
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
        ///     Return 1 (positive) if left greater than left
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
                        rightNumber => 1,
                        rightText => string.Compare(leftText.Value, rightText.Value, culture, CompareOptions.IgnoreCase),
                        rightError => rightError),
                leftError => leftError);
        }

        public static ScalarValue GetCellValue(XLRangeAddress area, int row, int column, CalcContext ctx)
        {
            var worksheet = area.Worksheet ?? ctx.Worksheet;
            var cell = worksheet.GetCell(row, column);
            if (cell is null)
                return Number1.Zero;

            if (cell.IsEvaluating)
                throw new InvalidOperationException("Formula has a circular reference");

            var value = cell.Value;
            // TODO: Replace with a conversion, like ctx
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

        /// <summary>
        /// Convert any kind of formula value to value returned as a content of a cell.
        /// <list type="bullet">
        ///    <item><c>bool</c> - represents a logical value.</item>
        ///    <item><c>double</c> - represents a number and also date/time as serial date-time.</item>
        ///    <item><c>string</c> - represents a text value.</item>
        ///    <item><see cref="ExpressionErrorType" /> - represents a formula calculation error.</item>
        /// </list>
        /// </summary>
        public static object ToCellContentValue(this AnyValue value, CalcContext ctx)
        {
            if (value.TryPickScalar(out var scalar, out var collection))
                return scalar.ToCellContentValue();

            return collection.Match(
                array => array[0, 0].ToCellContentValue(),
                reference =>
                {
                    if (reference.IsSingleCell())
                    {
                        // TODO: Better API
                        var a = reference.Areas.Single();
                        var cellValue = ctx.GetCellValue(a.Worksheet, a.FirstAddress.RowNumber, a.FirstAddress.ColumnNumber);
                        return cellValue.ToCellContentValue();
                    }

                    return reference
                        .ImplicitIntersection(ctx.FormulaAddress)
                        .Match<object>(
                            singleCellReference =>
                            {
                                var cellArea = singleCellReference.Areas.Single();
                                var cellAddress = cellArea.FirstAddress;
                                var cellValue = ctx.GetCellValue(cellArea.Worksheet, cellAddress.RowNumber, cellAddress.ColumnNumber);
                                return cellValue.ToCellContentValue();
                            },
                            error => error.Type);
                });
        }

        public static object ToCellContentValue(this ScalarValue value)
        {
            return value.Match<object>(
                logical => logical.Value,
                number => number.Value,
                text => text.Value,
                error => error.Type);
        }
    }
}
