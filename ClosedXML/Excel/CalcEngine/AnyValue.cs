using System;
using System.Globalization;
using System.Linq;
using CollectionValue = ClosedXML.Excel.CalcEngine.OneOf<ClosedXML.Excel.CalcEngine.Array, ClosedXML.Excel.CalcEngine.Reference>;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// A representation of any value that can be passed around in the formula evaluation.
    /// </summary>
    internal readonly struct AnyValue
    {
        private readonly byte _index;
        private readonly bool _logical;
        private readonly double _number;
        private readonly string _text;
        private readonly Error _error;
        private readonly Array _array;
        private readonly Reference _reference;

        private AnyValue(byte index, bool logical, double number, string text, Error error, Array array, Reference reference)
        {
            _index = index;
            _logical = logical;
            _number = number;
            _text = text;
            _error = error;
            _array = array;
            _reference = reference;
        }

        public static AnyValue From(bool logical) => new(0, logical, default, default, default, default, default);

        public static AnyValue From(double number) => new(1, default, number, default, default, default, default);

        public static AnyValue From(string text)
        {
            if (text is null)
                throw new ArgumentNullException();

            return new AnyValue(2, default, default, text, default, default, default);
        }

        public static AnyValue From(Error error) => new(3, default, default, default, error, default, default);

        public static AnyValue From(Array array)
        {
            if (array is null)
                throw new ArgumentNullException();

            return new(4, default, default, default, default, array, default);
        }

        public static AnyValue From(Reference reference)
        {
            if (reference is null)
                throw new ArgumentNullException();

            return new(5, default, default, default, default, default, reference);
        }

        public static implicit operator AnyValue(bool logical) => From(logical);

        public static implicit operator AnyValue(double number) => From(number);

        public static implicit operator AnyValue(string text) => From(text);

        public static implicit operator AnyValue(Error error) => From(error);

        public static implicit operator AnyValue(Array array) => From(array);

        public static implicit operator AnyValue(Reference reference) => From(reference);

        public bool TryPickScalar(out ScalarValue scalar, out CollectionValue collection)
        {
            scalar = _index switch
            {
                0 => _logical,
                1 => _number,
                2 => _text,
                3 => _error,
                _ => default
            };
            collection = _index switch
            {
                4 => _array,
                5 => _reference,
                _ => default
            };
            return _index <= 3;
        }

        public bool TryPickReference(out Reference reference)
        {
            if (_index == 5)
            {
                reference = _reference;
                return true;
            }

            reference = default;
            return false;
        }

        public TResult Match<TResult>(Func<bool, TResult> transformLogical, Func<double, TResult> transformNumber, Func<string, TResult> transformText, Func<Error, TResult> transformError, Func<Array, TResult> transformArray, Func<Reference, TResult> transformReference)
        {
            return _index switch
            {
                0 => transformLogical(_logical),
                1 => transformNumber(_number),
                2 => transformText(_text),
                3 => transformError(_error),
                4 => transformArray(_array),
                5 => transformReference(_reference),
                _ => throw new InvalidOperationException()
            };
        }

        #region Reference operators

        /// <summary>
        /// Implicit intersection for arguments of functions that don't accept range as a parameter (Excel 2016).
        /// </summary>
        /// <returns>Unchanged value for anything other than reference. Reference is changed into a single cell/#VALUE!</returns>
        public AnyValue ImplicitIntersection(CalcContext context)
        {
            return Match(
                logical => logical,
                number => number,
                text => text,
                logical => logical,
                array => array, // Array is unaffected by implicit intersection for operands - e.g. MMULT(COS({0,0});COS({0;0})) = 2
                reference =>
                {
                    if (reference.IsSingleCell())
                        return reference;

                    return reference
                        .ImplicitIntersection(context.FormulaAddress).Match<AnyValue>(
                            singleCellReference => singleCellReference,
                            error => error);
                });
        }

        /// <summary>
        /// Create a new reference that has one area that contains both operands or #VALUE! if not possible.
        /// </summary>
        public static AnyValue ReferenceRange(in AnyValue left, in AnyValue right, CalcContext ctx)
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
        public static AnyValue ReferenceUnion(in AnyValue left, in AnyValue right)
        {
            var leftConversionResult = ConvertToReference(left);
            if (!leftConversionResult.TryPickT0(out var leftReference, out var leftError))
                return leftError;

            var rightConversionResult = ConvertToReference(right);
            if (!rightConversionResult.TryPickT0(out var rightReference, out var rightError))
                return rightError;

            return Reference.UnionOp(leftReference, rightReference);
        }

        private static OneOf<Reference, Error> ConvertToReference(in AnyValue value)
        {
            return value.Match<OneOf<Reference, Error>>(
                logical => Error.CellValue,
                number => Error.CellValue,
                text => Error.CellValue,
                error => error,
                array => Error.CellValue,
                reference => reference);
        }

        #endregion

        #region Arithmetic unary operations

        public AnyValue UnaryPlus()
        {
            // Unary plus doesn't even convert to number. Type and value is retained.
            return this;
        }

        public AnyValue UnaryMinus(CalcContext context)
            => UnaryOperation(this, x => -x, context);

        public AnyValue UnaryPercent(CalcContext context)
            => UnaryOperation(this, x => x / 100.0, context);

        private static AnyValue UnaryOperation(in AnyValue value, Func<double, double> operatorFn, CalcContext context)
        {
            if (value.TryPickScalar(out var scalar, out var collection))
                return UnaryArithmeticOp(scalar, operatorFn, context.Converter).ToAnyValue();

            return collection.Match(
                array => array.Apply(arrayConst => UnaryArithmeticOp(arrayConst, operatorFn, context.Converter)),
                reference => reference
                    .Apply(cellValue => UnaryArithmeticOp(cellValue, operatorFn, context.Converter), context)
                    .Match<AnyValue>(
                        array => array,
                        error => error));
        }

        private static ScalarValue UnaryArithmeticOp(ScalarValue value, Func<double, double> op, ValueConverter converter)
        {
            var conversionResult = converter.CovertToNumber(value);
            if (!conversionResult.TryPickT0(out var number, out var error))
                return error;

            return op(number);
        }

        #endregion

        #region Arithmetic binary operators

        public static AnyValue BinaryPlus(in AnyValue left, in AnyValue right, CalcContext context)
        {
            return BinaryOperation(left, right, Plus, context);

            ScalarValue Plus(ScalarValue leftItem, ScalarValue rightItem)
            {
                return BinaryArithmeticOp(leftItem, rightItem, (lhs, rhs) => lhs + rhs, context.Converter);
            }
        }

        public static AnyValue BinaryMinus(in AnyValue left, in AnyValue right, CalcContext context)
        {
            return BinaryOperation(left, right, Minus, context);

            ScalarValue Minus(ScalarValue leftItem, ScalarValue rightItem)
            {
                return BinaryArithmeticOp(leftItem, rightItem, (lhs, rhs) => lhs - rhs, context.Converter);
            }
        }

        public static AnyValue BinaryMult(in AnyValue left, in AnyValue right, CalcContext context)
        {
            return BinaryOperation(left, right, Mult, context);

            ScalarValue Mult(ScalarValue leftItem, ScalarValue rightItem)
            {
                return BinaryArithmeticOp(leftItem, rightItem, (lhs, rhs) => lhs * rhs, context.Converter);
            }
        }

        public static AnyValue BinaryDiv(in AnyValue left, in AnyValue right, CalcContext context)
        {
            return BinaryOperation(left, right, Div, context);

            ScalarValue Div(ScalarValue leftItem, ScalarValue rightItem)
            {
                return BinaryArithmeticOp(leftItem, rightItem, (lhs, rhs) => rhs == 0.0 ? Error.DivisionByZero : lhs / rhs, context.Converter);
            }
        }

        public static AnyValue BinaryExp(in AnyValue left, in AnyValue right, CalcContext context)
        {
            return BinaryOperation(left, right, Exp, context);

            ScalarValue Exp(ScalarValue leftItem, ScalarValue rightItem)
            {
                return BinaryArithmeticOp(leftItem, rightItem, (lhs, rhs) => lhs == 0 && rhs == 0 ? Error.NumberInvalid : Math.Pow(lhs, rhs), context.Converter);
            }
        }

        #endregion

        #region Comparison operators

        public static AnyValue IsEqual(in AnyValue left, in AnyValue right, CalcContext context)
        {
            return BinaryOperation(left, right, IsEqual, context);

            ScalarValue IsEqual(ScalarValue leftItem, ScalarValue rightItem)
            {
                return CompareValues(leftItem, rightItem, context.Culture).Match<ScalarValue>(
                    cmp => cmp == 0,
                    error => error);
            }
        }

        public static AnyValue IsNotEqual(in AnyValue left, in AnyValue right, CalcContext context)
        {
            return BinaryOperation(left, right, IsNotEqual, context);

            ScalarValue IsNotEqual(ScalarValue leftItem, ScalarValue rightItem)
            {
                return CompareValues(leftItem, rightItem, context.Culture).Match<ScalarValue>(
                    cmp => cmp != 0,
                    error => error);
            }
        }

        public static AnyValue IsGreaterThan(in AnyValue left, in AnyValue right, CalcContext context)
        {
            return BinaryOperation(left, right, IsGreaterThan, context);

            ScalarValue IsGreaterThan(ScalarValue leftItem, ScalarValue rightItem)
            {
                return CompareValues(leftItem, rightItem, context.Culture).Match<ScalarValue>(
                    cmp => cmp > 0,
                    error => error);
            }
        }

        public static AnyValue IsGreaterThanOrEqual(in AnyValue left, in AnyValue right, CalcContext context)
        {
            return BinaryOperation(left, right, IsGreaterThanOrEqual, context);

            ScalarValue IsGreaterThanOrEqual(ScalarValue leftItem, ScalarValue rightItem)
            {
                return CompareValues(leftItem, rightItem, context.Culture).Match<ScalarValue>(
                    cmp => cmp >= 0,
                    error => error);
            }
        }

        public static AnyValue IsLessThan(in AnyValue left, in AnyValue right, CalcContext context)
        {
            return BinaryOperation(left, right, IsLessThan, context);

            ScalarValue IsLessThan(ScalarValue leftItem, ScalarValue rightItem)
            {
                return CompareValues(leftItem, rightItem, context.Culture).Match<ScalarValue>(
                    cmp => cmp < 0,
                    error => error);
            }
        }

        public static AnyValue IsLessThanOrEqual(in AnyValue left, in AnyValue right, CalcContext context)
        {
            return BinaryOperation(left, right, IsLessThanOrEqual, context);

            ScalarValue IsLessThanOrEqual(ScalarValue leftItem, ScalarValue rightItem)
            {
                return CompareValues(leftItem, rightItem, context.Culture).Match<ScalarValue>(
                    cmp => cmp <= 0,
                    error => error);
            }
        }

        #endregion

        public static AnyValue Concat(in AnyValue left, in AnyValue right, CalcContext context)
        {
            return BinaryOperation(left, right, Concat, context);

            ScalarValue Concat(ScalarValue leftItem, ScalarValue rightItem)
            {
                return context.Converter.ToText(leftItem)
                    .Match(
                        leftText => context.Converter.ToText(rightItem).Match<OneOf<string, Error>>(
                            rightText => leftText + rightText,
                            rightError => rightError),
                        leftError => leftError).Match<ScalarValue>(
                            text => text,
                            error => error);
            }
        }

        private static AnyValue BinaryOperation(in AnyValue left, in AnyValue right, BinaryFunc func, CalcContext context)
        {
            var isLeftScalar = left.TryPickScalar(out var leftScalar, out var leftCollection);
            var isRightScalar = right.TryPickScalar(out var rightScalar, out var rightCollection);

            if (isLeftScalar && isRightScalar)
                return func(leftScalar, rightScalar).ToAnyValue();

            if (isLeftScalar)
            {
                return rightCollection.Match(
                    array => new ScalarArray(leftScalar, array.Width, array.Height).Apply(array, func),
                    rightReference =>
                    {
                        if (rightReference.TryGetSingleCellValue(out var rightCellValue, context))
                            return func(leftScalar, rightCellValue).ToAnyValue();

                        var referenceArrayResult = rightReference.ToArray(context);
                        if (!referenceArrayResult.TryPickT0(out var rightRefArray, out var rightError))
                            return rightError;

                        return new ScalarArray(leftScalar, rightRefArray.Width, rightRefArray.Height).Apply(rightRefArray, func);
                    });
            }

            if (isRightScalar)
            {
                return leftCollection.Match(
                    leftArray => leftArray.Apply(new ScalarArray(rightScalar, leftArray.Width, leftArray.Height), func),
                    leftReference =>
                    {
                        if (leftReference.TryGetSingleCellValue(out var leftCellValue, context))
                            return func(leftCellValue, rightScalar).ToAnyValue();

                        var referenceArrayResult = leftReference.ToArray(context);
                        if (!referenceArrayResult.TryPickT0(out var leftRefArray, out var leftError))
                            return leftError;

                        return leftRefArray.Apply(new ScalarArray(rightScalar, leftRefArray.Width, leftRefArray.Height), func);
                    });
            }

            // Both are aggregates
            return leftCollection.Match(
                leftArray => rightCollection.Match(
                        rightArray =>
                        {
                            var width = Math.Max(leftArray.Width, rightArray.Width);
                            var height = Math.Max(leftArray.Height, rightArray.Height);
                            return new ResizedArray(leftArray, width, height).Apply(new ResizedArray(rightArray, width, height), func);
                        },
                        rightReference =>
                        {
                            if (rightReference.TryGetSingleCellValue(out var rightCellValue, context))
                                return leftArray.Apply(new ScalarArray(rightCellValue, leftArray.Width, leftArray.Height), func);

                            if (rightReference.Areas.Count == 1)
                            {
                                var area = rightReference.Areas[0];
                                var width = Math.Max(leftArray.Width, area.ColumnSpan);
                                var height = Math.Max(leftArray.Height, area.RowSpan);
                                return new ResizedArray(leftArray, width, height).Apply(new ResizedArray(new ReferenceArray(area, context), width, height), func);
                            }

                            return leftArray.Apply(new ScalarArray(Error.CellValue, leftArray.Width, leftArray.Height), func);
                        }),
                leftReference => rightCollection.Match(
                        rightArray =>
                        {
                            if (leftReference.TryGetSingleCellValue(out var leftCellValue, context))
                                return new ScalarArray(leftCellValue, rightArray.Width, rightArray.Height).Apply(rightArray, func);

                            if (leftReference.Areas.Count == 1)
                            {
                                var area = leftReference.Areas[0];
                                var width = Math.Max(area.ColumnSpan, rightArray.Width);
                                var height = Math.Max(area.RowSpan, rightArray.Height);
                                var leftRefArray = new ResizedArray(new ReferenceArray(area, context), width, height);
                                return leftRefArray.Apply(new ResizedArray(rightArray, width, height), func);
                            }

                            var errorArray = new ScalarArray(Error.CellValue, rightArray.Width, rightArray.Height);
                            return errorArray.Apply(rightArray, func);
                        },
                        rightReference =>
                        {
                            if (leftReference.Areas.Count > 1 && rightReference.Areas.Count > 1)
                                return Error.CellValue;

                            if (leftReference.Areas.Count > 1)
                                return new ScalarArray(Error.CellValue, rightReference.Areas[0].ColumnSpan, rightReference.Areas[0].RowSpan);

                            if (rightReference.Areas.Count > 1)
                                return new ScalarArray(Error.CellValue, leftReference.Areas[0].ColumnSpan, leftReference.Areas[0].RowSpan);

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

                            var leftRefArray = new ResizedArray(new ReferenceArray(leftArea, context), colSpan, rowSpan);
                            var rightRefArray = new ResizedArray(new ReferenceArray(rightArea, context), colSpan, rowSpan);
                            return leftRefArray.Apply(rightRefArray, func);
                        }));
        }

        private static ScalarValue BinaryArithmeticOp(ScalarValue left, ScalarValue right, BinaryNumberFunc func, ValueConverter converter)
        {
            var leftConversionResult = converter.CovertToNumber(left);
            if (!leftConversionResult.TryPickT0(out var leftNumber, out var leftError))
            {
                return leftError;
            }

            var rightConversionResult = converter.CovertToNumber(right);
            if (!rightConversionResult.TryPickT0(out var rightNumber, out var rightError))
            {
                return rightError;
            }

            return func(leftNumber, rightNumber).Match<ScalarValue>(
                number => number,
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
            return left.Match(culture,
                (leftLogical, _) => right.Match<OneOf<int, Error>, bool>(leftLogical,
                        (rightLogical, leftLogical) => leftLogical.CompareTo(rightLogical),
                        (rightNumber, _) => 1,
                        (rightText, _) => 1,
                        (rightError, _) => rightError),
                (leftNumber, _) => right.Match<OneOf<int, Error>, double>(leftNumber,
                        (rightLogical, _) => -1,
                        (rightNumber, leftNumber) => leftNumber.CompareTo(rightNumber),
                        (rightText, _) => -1,
                        (rightError, _) => rightError),
                (leftText, culture) => right.Match<OneOf<int, Error>, string, CultureInfo>(leftText, culture,
                        (rightLogical, _, _) => -1,
                        (rightNumber, _, _) => 1,
                        (rightText, leftText, culture) => string.Compare(leftText, rightText, culture, CompareOptions.IgnoreCase),
                        (rightError, _, _) => rightError),
                (leftError, _) => leftError);
        }

        private delegate OneOf<double, Error> BinaryNumberFunc(double lhs, double rhs);
    }

    internal delegate ScalarValue BinaryFunc(ScalarValue lhs, ScalarValue rhs);
}
