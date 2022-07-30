using OneOf;
using System;
using System.Globalization;
using System.Linq;
using AnyValue = OneOf.OneOf<bool, double, string, ClosedXML.Excel.CalcEngine.Error, ClosedXML.Excel.CalcEngine.Array, ClosedXML.Excel.CalcEngine.Reference>;
using ScalarValue = OneOf.OneOf<bool, double, string, ClosedXML.Excel.CalcEngine.Error>;
using CollectionValue = OneOf.OneOf<ClosedXML.Excel.CalcEngine.Array, ClosedXML.Excel.CalcEngine.Reference>;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// Implementation of reference, arithmetic, text and comparison operators on AnyValue that use Excel semantic.
    /// </summary>
    internal static class OperatorExtensions
    {
        #region Reference operators

        /// <summary>
        /// Implicit intersection for arguments of functions that don't accept range as a parameter (Excel 2016).
        /// </summary>
        /// <returns>Unchanged value for anything other than reference. Reference is changed into a single cell/#VALUE!</returns>
        public static AnyValue ImplicitIntersection(this AnyValue value, CalcContext context)
        {
            return value.Match(
                logical => logical,
                number => number,
                text => text,
                logical => logical,
                array => array, // Yup, array is unaffected by implicit intersection for operands - MMULT(COS({0,0});COS({0;0})) = 2
                reference =>
                {
                    if (reference.IsSingleCell())
                        return reference;

                    return reference
                        .ImplicitIntersection(context.FormulaAddress)
                        .Match<AnyValue>(
                            singleCellReference => singleCellReference,
                            error => error);
                });
        }

        /// <summary>
        /// Create a new reference that has one area that contains both operands.
        /// </summary>
        public static AnyValue ReferenceRange(this AnyValue left, AnyValue right, CalcContext ctx)
        {
            var leftConversionResult = ConvertToReference(left);
            if (!leftConversionResult.TryPickT0(out var leftReference, out var leftError))
                return leftError;

            var rightConversionResult = ConvertToReference(right);
            if (!rightConversionResult.TryPickT0(out var rightReference, out var rightError))
                return rightError;

            return Reference.RangeOp(leftReference, rightReference, ctx.Worksheet).Match<AnyValue>(
                reference => reference,
                error => error);
        }

        /// <summary>
        /// Create a new reference by combining areas of both arguments. Areas of the new reference can overlap = some overlapping
        /// cells might be counted multiple times (<c>SUM((A1;A1)) = 2</c> if <c>A1</c> is <c>1</c>).
        /// </summary>
        public static AnyValue ReferenceUnion(this AnyValue left, AnyValue right)
        {
            var leftConversionResult = ConvertToReference(left);
            if (!leftConversionResult.TryPickT0(out var leftReference, out var leftError))
                return leftError;

            var rightConversionResult = ConvertToReference(right);
            if (!rightConversionResult.TryPickT0(out var rightReference, out var rightError))
                return rightError;

            return Reference.UnionOp(leftReference, rightReference);
        }

        private static OneOf<Reference, Error> ConvertToReference(AnyValue left)
        {
            return left.Match<OneOf<Reference, Error>>(
                logical => Error.CellValue,
                number => Error.CellValue,
                text => Error.CellValue,
                error => error,
                array => Error.CellValue,
                reference => reference);
        }

        #endregion

        #region Arithmetic unary operations

        public static AnyValue UnaryPlus(this AnyValue value)
        {
            // Unary plus doesn't even convert to number. Type and value is retained.
            return value;
        }

        public static AnyValue UnaryMinus(this AnyValue value, CalcContext context)
            => UnaryOperation(value, x => -x, context);

        public static AnyValue UnaryPercent(this AnyValue value, CalcContext context)
            => UnaryOperation(value, x => x / 100.0, context);

        private static AnyValue UnaryOperation(AnyValue value, Func<double, double> operatorFn, CalcContext context)
        {
            if (value.TryPickScalar(out var scalar, out var collection))
                return UnaryArithmeticOp(scalar, operatorFn, context.Converter).ToAnyValue();

            return collection.Match(
                array => ApplyOnArray(array, arrayConst => UnaryArithmeticOp(arrayConst, operatorFn, context.Converter)),
                reference => ApplyOnReference(reference, cellValue => UnaryArithmeticOp(cellValue, operatorFn, context.Converter), context));
        }

        private static AnyValue ApplyOnArray(Array array, Func<ScalarValue, ScalarValue> op)
        {
            var data = new ScalarValue[array.Height, array.Width];
            for (int y = 0; y < array.Height; ++y)
                for (int x = 0; x < array.Width; ++x)
                    data[y, x] = op(array[y, x]);

            return new ConstArray(data);
        }

        private static AnyValue ApplyOnReference(Reference reference, Func<ScalarValue, ScalarValue> op, CalcContext context)
        {
            if (reference.Areas.Count != 1)
                return Error.CellValue;

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
                    var cellValue = context.GetCellValue(area.Worksheet, row, column);
                    data[y, x] = op(cellValue);
                }
            }
            return new ConstArray(data);
        }

        private static ScalarValue UnaryArithmeticOp(ScalarValue value, Func<double, double> op, ValueConverter converter)
        {
            var conversionResult = CovertToNumber(value, converter);
            if (!conversionResult.TryPickT0(out var number, out var error))
                return error;

            return op(number);
        }

        #endregion

        #region Arithmetic binary operators

        public static AnyValue BinaryPlus(this AnyValue left, AnyValue right, CalcContext context)
        {
            return BinaryOperation(left, right, Plus, context);

            ScalarValue Plus(ScalarValue leftItem, ScalarValue rightItem)
                => BinaryArithmeticOp(leftItem, rightItem, (lhs, rhs) => lhs + rhs, context.Converter);
        }

        public static AnyValue BinaryMinus(this AnyValue left, AnyValue right, CalcContext context)
        {
            return BinaryOperation(left, right, Minus, context);

            ScalarValue Minus(ScalarValue leftItem, ScalarValue rightItem)
                => BinaryArithmeticOp(leftItem, rightItem, (lhs, rhs) => lhs - rhs, context.Converter);
        }

        public static AnyValue BinaryMult(this AnyValue left, AnyValue right, CalcContext context)
        {
            return BinaryOperation(left, right, Mult, context);

            ScalarValue Mult(ScalarValue leftItem, ScalarValue rightItem)
                => BinaryArithmeticOp(leftItem, rightItem, (lhs, rhs) => lhs * rhs, context.Converter);
        }

        public static AnyValue BinaryDiv(this AnyValue left, AnyValue right, CalcContext context)
        {
            return BinaryOperation(left, right, Div, context);

            ScalarValue Div(ScalarValue leftItem, ScalarValue rightItem)
                => BinaryArithmeticOp(leftItem, rightItem, (lhs, rhs) => rhs == 0.0 ? Error.DivisionByZero : lhs / rhs, context.Converter);
        }

        public static AnyValue BinaryExp(this AnyValue left, AnyValue right, CalcContext context)
        {
            return BinaryOperation(left, right, Exp, context);

            ScalarValue Exp(ScalarValue leftItem, ScalarValue rightItem)
                => BinaryArithmeticOp(leftItem, rightItem, (lhs, rhs) => lhs == 0 && rhs == 0 ? Error.NumberInvalid : Math.Pow(lhs, rhs), context.Converter);
        }

        #endregion

        #region Comparison operators

        public static AnyValue IsEqual(this AnyValue left, AnyValue right, CalcContext context)
        {
            BinaryFunc g = (leftItem, rightItem) => CompareValues(leftItem, rightItem, context.Culture).Match<ScalarValue>(
                cmp => cmp == 0,
                error => error);
            return BinaryOperation(left, right, g, context);
        }

        public static AnyValue IsNotEqual(this AnyValue left, AnyValue right, CalcContext context)
        {
            BinaryFunc g = (leftItem, rightItem) => CompareValues(leftItem, rightItem, context.Culture).Match<ScalarValue>(
                cmp => cmp != 0,
                error => error);
            return BinaryOperation(left, right, g, context);
        }

        public static AnyValue IsGreaterThan(this AnyValue left, AnyValue right, CalcContext context)
        {
            BinaryFunc g = (leftItem, rightItem) => CompareValues(leftItem, rightItem, context.Culture).Match<ScalarValue>(
                cmp => cmp > 0,
                error => error);
            return BinaryOperation(left, right, g, context);
        }

        public static AnyValue IsGreaterThanOrEqual(this AnyValue left, AnyValue right, CalcContext context)
        {
            BinaryFunc g = (leftItem, rightItem) => CompareValues(leftItem, rightItem, context.Culture).Match<ScalarValue>(
                cmp => cmp >= 0,
                error => error);
            return BinaryOperation(left, right, g, context);
        }

        public static AnyValue IsLessThan(this AnyValue left, AnyValue right, CalcContext context)
        {
            BinaryFunc g = (leftItem, rightItem) => CompareValues(leftItem, rightItem, context.Culture).Match<ScalarValue>(
                cmp => cmp < 0,
                error => error);
            return BinaryOperation(left, right, g, context);
        }

        public static AnyValue IsLessThanOrEqual(this AnyValue left, AnyValue right, CalcContext context)
        {
            BinaryFunc g = (leftItem, rightItem) => CompareValues(leftItem, rightItem, context.Culture).Match<ScalarValue>(
                cmp => cmp <= 0,
                error => error);
            return BinaryOperation(left, right, g, context);
        }

        #endregion

        public static AnyValue Concat(this AnyValue left, AnyValue right, CalcContext context)
        {
            BinaryFunc g = (lhs, rhs) =>
            {
                return context.Converter.ToText(lhs).Match(
                    leftText => context.Converter.ToText(rhs).Match<OneOf<string, Error>>(rightText => leftText + rightText, rightError => rightError),
                    leftError => leftError).Match<ScalarValue>(text => text, error => error);
            };

            return BinaryOperation(left, right, g, context);
        }

        private static AnyValue BinaryOperation(AnyValue left, AnyValue right, BinaryFunc func, CalcContext context)
        {
            var isLeftScalar = left.TryPickScalar(out var leftScalar, out var leftCollection);
            var isRightScalar = right.TryPickScalar(out var rightScalar, out var rightCollection);

            if (isLeftScalar && isRightScalar)
                return func(leftScalar, rightScalar).ToAnyValue();

            if (isLeftScalar)
            {
                return rightCollection.Match(
                    array => ApplyOnArray(
                        new ScalarArray(leftScalar, array.Width, array.Height),
                        array,
                        func),
                    rightReference =>
                    {
                        if (rightReference.TryGetSingleCellValue(out var rightCellValue, context))
                            return func(leftScalar, rightCellValue).ToAnyValue();

                        var referenceArrayResult = rightReference.ToArray(context);
                        if (!referenceArrayResult.TryPickT0(out var rightRefArray, out var rightError))
                            return rightError;

                        return ApplyOnArray(new ScalarArray(leftScalar, rightRefArray.Width, rightRefArray.Height), rightRefArray, func);
                    });
            }

            if (isRightScalar)
            {
                return leftCollection.Match(
                    leftArray => ApplyOnArray(
                        leftArray,
                        new ScalarArray(rightScalar, leftArray.Width, leftArray.Height),
                        func),
                    leftReference =>
                    {
                        if (leftReference.TryGetSingleCellValue(out var leftCellValue, context))
                            return func(leftCellValue, rightScalar).ToAnyValue();

                        var referenceArrayResult = leftReference.ToArray(context);
                        if (!referenceArrayResult.TryPickT0(out var leftRefArray, out var leftError))
                            return leftError;

                        return ApplyOnArray(leftRefArray, new ScalarArray(rightScalar, leftRefArray.Width, leftRefArray.Height), func);
                    });
            }

            // Both are aggregates
            return leftCollection.Match(
                leftArray => rightCollection.Match(
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
                leftReference => rightCollection.Match(
                        rightArray =>
                        {
                            throw new NotImplementedException();
                        },
                        rightReference =>
                        {
                            if (leftReference.Areas.Count > 1 || rightReference.Areas.Count > 1)
                                return Error.CellValue;

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

        private static OneOf<Array, Error> ToArray(this Reference reference, CalcContext context)
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
            return new ConstArray(data);
        }

        private static ScalarValue BinaryArithmeticOp(ScalarValue left, ScalarValue right, BinaryNumberFunc func, ValueConverter converter)
        {
            var leftConversionResult = left.CovertToNumber(converter);
            if (!leftConversionResult.TryPickT0(out var leftNumber, out var leftError))
            {
                return leftError;
            }

            var rightConversionResult = right.CovertToNumber(converter);
            if (!rightConversionResult.TryPickT0(out var rightNumber, out var rightError))
            {
                return rightError;
            }

            return func(leftNumber, rightNumber).Match<ScalarValue>(
                number => number,
                error => error);
        }

        private static OneOf<double, Error> CovertToNumber(this ScalarValue value, ValueConverter converter)
        {
            return value.Match(
                logical => logical ? 1 : 0,
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
        /// <param name="left">Left hand operand of the comparison.</param>
        /// <param name="right">Right hand operand of the comparison.</param>
        /// <param name="culture">Culture to use for comparison.</param>
        /// <returns>
        ///     Return -1 (negative)  if left less than right
        ///     Return 1 (positive) if left greater than left
        ///     Return 0 if both operands are considered equal.
        /// </returns>
        private static OneOf<int, Error> CompareValues(ScalarValue left, ScalarValue right, CultureInfo culture)
        {
            return left.Match(
                leftLogical => right.Match<OneOf<int, Error>>(
                        rightLogical => leftLogical.CompareTo(rightLogical),
                        rightNumber => 1,
                        rightText => 1,
                        rightError => rightError),
                leftNumber => right.Match<OneOf<int, Error>>(
                        rightLogical => -1,
                        rightNumber => leftNumber.CompareTo(rightNumber),
                        rightText => -1,
                        rightError => rightError),
                leftText => right.Match<OneOf<int, Error>>(
                        rightLogical => -1,
                        rightNumber => 1,
                        rightText => string.Compare(leftText, rightText, culture, CompareOptions.IgnoreCase),
                        rightError => rightError),
                leftError => leftError);
        }

        private delegate ScalarValue BinaryFunc(ScalarValue lhs, ScalarValue rhs);

        private delegate OneOf<double, Error> BinaryNumberFunc(double lhs, double rhs);

        #region Type conversion functions

        public static bool TryPickScalar(this AnyValue value, out ScalarValue scalar, out CollectionValue collection)
        {
            scalar = value.Index switch
            {
                0 => value.AsT0,
                1 => value.AsT1,
                2 => value.AsT2,
                3 => value.AsT3,
                _ => default
            };
            collection = value.Index switch
            {
                4 => value.AsT4,
                5 => value.AsT5,
                _ => default
            };
            return value.Index <= 3;
        }

        public static AnyValue ToAnyValue(this ScalarValue scalar)
        {
            return scalar.Match<AnyValue>(
                logical => logical,
                number => number,
                text => text,
                error => error);
        }

        /// <summary>
        /// Convert any kind of formula value to value returned as a content of a cell.
        /// <list type="bullet">
        ///    <item><c>bool</c> - represents a logical value.</item>
        ///    <item><c>double</c> - represents a number and also date/time as serial date-time.</item>
        ///    <item><c>string</c> - represents a text value.</item>
        ///    <item><see cref="Error" /> - represents a formula calculation error.</item>
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
                    if (reference.TryGetSingleCellValue(out var cellValue, ctx))
                        return cellValue.ToCellContentValue();

                    return reference
                        .ImplicitIntersection(ctx.FormulaAddress)
                        .Match<object>(
                            singleCellReference =>
                            {
                                if (!singleCellReference.TryGetSingleCellValue(out var cellValue, ctx))
                                    throw new InvalidOperationException();

                                return cellValue.ToCellContentValue();
                            },
                            error => error);
                });
        }

        public static object ToCellContentValue(this ScalarValue value)
        {
            return value.Match<object>(
                logical => logical,
                number => number,
                text => text,
                error => error);
        }

        #endregion
    }
}
