using System;
using System.Globalization;
using CollectionValue = ClosedXML.Excel.CalcEngine.OneOf<ClosedXML.Excel.CalcEngine.Array, ClosedXML.Excel.CalcEngine.Reference>;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// A discriminated union representing any value that can be passed around in the formula evaluation.
    /// </summary>
    internal readonly struct AnyValue
    {
        private const int BlankValue = 0;
        private const int LogicalValue = 1;
        private const int NumberValue = 2;
        private const int TextValue = 3;
        private const int ErrorValue = 4;
        private const int ArrayValue = 5;
        private const int ReferenceValue = 6;

        private readonly byte _index;
        private readonly bool _logical;
        private readonly double _number;
        private readonly string _text;
        private readonly XLError _error;
        private readonly Array _array;
        private readonly Reference _reference;

        private AnyValue(byte index, bool logical, double number, string text, XLError error, Array array, Reference reference)
        {
            _index = index;
            _logical = logical;
            _number = number;
            _text = text;
            _error = error;
            _array = array;
            _reference = reference;
        }

        /// <summary>
        /// A value of a blank cell or missing argument. Conversion methods mostly treat blank like 0 or an empty string.
        /// </summary>
        public static readonly AnyValue Blank = new(BlankValue, default, default, default, default, default, default);

        public static AnyValue From(bool logical) => new(LogicalValue, logical, default, default, default, default, default);

        public static AnyValue From(double number) => new(NumberValue, default, number, default, default, default, default);

        public static AnyValue From(string text)
        {
            if (text is null)
                throw new ArgumentNullException();

            return new AnyValue(TextValue, default, default, text, default, default, default);
        }

        public static AnyValue From(XLError error) => new(ErrorValue, default, default, default, error, default, default);

        public static AnyValue From(Array array)
        {
            if (array is null)
                throw new ArgumentNullException();

            return new(ArrayValue, default, default, default, default, array, default);
        }

        public static AnyValue From(Reference reference)
        {
            if (reference is null)
                throw new ArgumentNullException();

            return new(ReferenceValue, default, default, default, default, default, reference);
        }

        public static implicit operator AnyValue(bool logical) => From(logical);

        public static implicit operator AnyValue(double number) => From(number);

        public static implicit operator AnyValue(string text) => From(text);

        public static implicit operator AnyValue(XLError error) => From(error);

        public static implicit operator AnyValue(Array array) => From(array);

        public static implicit operator AnyValue(Reference reference) => From(reference);

        public bool IsBlank => _index == BlankValue;

        public bool TryPickScalar(out ScalarValue scalar, out CollectionValue collection)
        {
            scalar = _index switch
            {
                BlankValue => ScalarValue.Blank,
                LogicalValue => _logical,
                NumberValue => _number,
                TextValue => _text,
                ErrorValue => _error,
                _ => default
            };
            collection = _index switch
            {
                ArrayValue => _array,
                ReferenceValue => _reference,
                _ => default
            };
            return _index <= ErrorValue;
        }

        public bool TryPickError(out XLError error)
        {
            if (_index == ErrorValue)
            {
                error = _error;
                return true;
            }

            error = default;
            return false;
        }

        public bool TryPickReference(out Reference reference, out XLError error)
        {
            if (_index == ReferenceValue)
            {
                reference = _reference;
                error = default;
                return true;
            }

            reference = default;
            error = _index == ErrorValue ? _error : XLError.IncompatibleValue;
            return false;
        }

        /// <summary>
        /// Try to get a reference that is a one area from the value.
        /// </summary>
        /// <param name="area">The found area.</param>
        /// <param name="error">Original error, if the value is error, <c>#VALUE!</c> if type is not a reference or #REF! if more than one area in the reference.</param>
        /// <returns>True if area could be determined, false otherwise.</returns>
        public bool TryPickArea(out XLRangeAddress area, out XLError error)
        {
            if (_index != ReferenceValue)
            {
                area = default;
                error = _index == ErrorValue ? _error : XLError.IncompatibleValue;
                return false;
            }

            if (_reference.Areas.Count > 1)
            {
                area = default;
                error = XLError.CellReference;
                return false;
            }

            area = _reference.Areas[0];
            error = default;
            return true;
        }

        public TResult Match<TResult>(Func<TResult> transformBlank, Func<bool, TResult> transformLogical, Func<double, TResult> transformNumber, Func<string, TResult> transformText, Func<XLError, TResult> transformError, Func<Array, TResult> transformArray, Func<Reference, TResult> transformReference)
        {
            return _index switch
            {
                BlankValue => transformBlank(),
                LogicalValue => transformLogical(_logical),
                NumberValue => transformNumber(_number),
                TextValue => transformText(_text),
                ErrorValue => transformError(_error),
                ArrayValue => transformArray(_array),
                ReferenceValue => transformReference(_reference),
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
                () => Blank,
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

        private static OneOf<Reference, XLError> ConvertToReference(in AnyValue value)
        {
            return value._index switch
            {
                ReferenceValue => value._reference,
                ErrorValue => value._error,
                _ => XLError.IncompatibleValue
            };
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
                return UnaryArithmeticOp(scalar, operatorFn, context).ToAnyValue();

            return collection.Match(
                array => array.Apply(arrayConst => UnaryArithmeticOp(arrayConst, operatorFn, context)),
                reference => reference
                    .Apply(cellValue => UnaryArithmeticOp(cellValue, operatorFn, context), context)
                    .Match<AnyValue>(
                        array => array,
                        error => error));
        }

        private static ScalarValue UnaryArithmeticOp(ScalarValue value, Func<double, double> op, CalcContext ctx)
        {
            var conversionResult = value.ToNumber(ctx.Culture);
            if (!conversionResult.TryPickT0(out var number, out var error))
                return error;

            return op(number);
        }

        #endregion

        #region Arithmetic binary operators

        public static AnyValue BinaryPlus(in AnyValue left, in AnyValue right, CalcContext context)
        {
            return BinaryOperation(in left, in right, static (in ScalarValue leftItem, in ScalarValue rightItem, CalcContext ctx) =>
            {
                return BinaryArithmeticOp(in leftItem, in rightItem, static (lhs, rhs) => lhs + rhs, ctx);
            }, context);
        }

        public static AnyValue BinaryMinus(in AnyValue left, in AnyValue right, CalcContext context)
        {
            return BinaryOperation(in left, in right, static (in ScalarValue leftItem, in ScalarValue rightItem, CalcContext ctx) =>
            {
                return BinaryArithmeticOp(in leftItem, in rightItem, static (lhs, rhs) => lhs - rhs, ctx);
            }, context);
        }

        public static AnyValue BinaryMult(in AnyValue left, in AnyValue right, CalcContext context)
        {
            return BinaryOperation(in left, in right, static (in ScalarValue leftItem, in ScalarValue rightItem, CalcContext ctx) =>
            {
                return BinaryArithmeticOp(in leftItem, in rightItem, static (lhs, rhs) => lhs * rhs, ctx);
            }, context);
        }

        public static AnyValue BinaryDiv(in AnyValue left, in AnyValue right, CalcContext context)
        {
            return BinaryOperation(in left, in right, static (in ScalarValue leftItem, in ScalarValue rightItem, CalcContext ctx) =>
            {
                return BinaryArithmeticOp(in leftItem, in rightItem, static (lhs, rhs) => rhs == 0.0 ? XLError.DivisionByZero : lhs / rhs, ctx);
            }, context);
        }

        public static AnyValue BinaryExp(in AnyValue left, in AnyValue right, CalcContext context)
        {
            return BinaryOperation(in left, in right, static (in ScalarValue leftItem, in ScalarValue rightItem, CalcContext ctx) =>
            {
                return BinaryArithmeticOp(in leftItem, in rightItem, static (lhs, rhs) => lhs == 0 && rhs == 0 ? XLError.NumberInvalid : Math.Pow(lhs, rhs), ctx);
            }, context);
        }

        #endregion

        #region Comparison operators

        public static AnyValue IsEqual(in AnyValue left, in AnyValue right, CalcContext context)
        {
            return BinaryOperation(in left, in right, static (in ScalarValue leftItem, in ScalarValue rightItem, CalcContext ctx) =>
            {
                return CompareValues(leftItem, rightItem, ctx.Culture).Match<ScalarValue>(
                    static cmp => cmp == 0,
                    static error => error);
            }, context);
        }

        public static AnyValue IsNotEqual(in AnyValue left, in AnyValue right, CalcContext context)
        {
            return BinaryOperation(in left, in right, static (in ScalarValue leftItem, in ScalarValue rightItem, CalcContext ctx) =>
            {
                return CompareValues(leftItem, rightItem, ctx.Culture).Match<ScalarValue>(
                    static cmp => cmp != 0,
                    static error => error);
            }, context);
        }

        public static AnyValue IsGreaterThan(in AnyValue left, in AnyValue right, CalcContext context)
        {
            return BinaryOperation(in left, in right, static (in ScalarValue leftItem, in ScalarValue rightItem, CalcContext ctx) =>
            {
                return CompareValues(leftItem, rightItem, ctx.Culture).Match<ScalarValue>(
                    static cmp => cmp > 0,
                    static error => error);
            }, context);
        }

        public static AnyValue IsGreaterThanOrEqual(in AnyValue left, in AnyValue right, CalcContext context)
        {
            return BinaryOperation(in left, in right, static (in ScalarValue leftItem, in ScalarValue rightItem, CalcContext ctx) =>
            {
                return CompareValues(leftItem, rightItem, ctx.Culture).Match<ScalarValue>(
                    static cmp => cmp >= 0,
                    static error => error);
            }, context);
        }

        public static AnyValue IsLessThan(in AnyValue left, in AnyValue right, CalcContext context)
        {
            return BinaryOperation(in left, in right, static (in ScalarValue leftItem, in ScalarValue rightItem, CalcContext ctx) =>
            {
                return CompareValues(leftItem, rightItem, ctx.Culture).Match<ScalarValue>(
                    static cmp => cmp < 0,
                    static error => error);
            }, context);
        }

        public static AnyValue IsLessThanOrEqual(in AnyValue left, in AnyValue right, CalcContext context)
        {
            return BinaryOperation(in left, in right, static (in ScalarValue leftItem, in ScalarValue rightItem, CalcContext ctx) =>
            {
                return CompareValues(leftItem, rightItem, ctx.Culture).Match<ScalarValue>(
                    static cmp => cmp <= 0,
                    static error => error);
            }, context);
        }

        #endregion

        public static AnyValue Concat(in AnyValue left, in AnyValue right, CalcContext context)
        {
            return BinaryOperation(in left, in right, static (in ScalarValue leftItem, in ScalarValue rightItem, CalcContext ctx) =>
            {
                var leftTextResult = leftItem.ToText(ctx.Culture);
                if (!leftTextResult.TryPickT0(out var leftText, out var leftError))
                    return leftError;

                var rightTextResult = rightItem.ToText(ctx.Culture);
                if (!rightTextResult.TryPickT0(out var rightText, out var rightError))
                    return rightError;

                return leftText + rightText;
            }, context);
        }

        private static AnyValue BinaryOperation(in AnyValue left, in AnyValue right, BinaryFunc func, CalcContext context)
        {
            var isLeftScalar = left.TryPickScalar(out var leftScalar, out var leftCollection);
            var isRightScalar = right.TryPickScalar(out var rightScalar, out var rightCollection);

            if (isLeftScalar && isRightScalar)
                return func(in leftScalar, in rightScalar, context).ToAnyValue();

            if (isLeftScalar)
            {
                // Right side is an array
                if (rightCollection.TryPickT0(out var rightArray, out var rightReference))
                    return new ScalarArray(leftScalar, rightArray.Width, rightArray.Height).Apply(rightArray, func, context);

                // Right side is a reference
                if (rightReference.TryGetSingleCellValue(out var rightCellValue, context))
                    return func(in leftScalar, in rightCellValue, context).ToAnyValue();

                var referenceArrayResult = rightReference.ToArray(context);
                if (!referenceArrayResult.TryPickT0(out var rightRefArray, out var rightError))
                    return rightError;

                return new ScalarArray(leftScalar, rightRefArray.Width, rightRefArray.Height).Apply(rightRefArray, func, context);
            }

            if (isRightScalar)
            {
                // Left side is an array
                if (leftCollection.TryPickT0(out var leftArray, out var leftReference))
                    return leftArray.Apply(new ScalarArray(rightScalar, leftArray.Width, leftArray.Height), func, context);

                // Left side is a reference
                if (leftReference.TryGetSingleCellValue(out var leftCellValue, context))
                    return func(leftCellValue, rightScalar, context).ToAnyValue();

                var referenceArrayResult = leftReference.ToArray(context);
                if (!referenceArrayResult.TryPickT0(out var leftRefArray, out var leftError))
                    return leftError;

                return leftRefArray.Apply(new ScalarArray(rightScalar, leftRefArray.Width, leftRefArray.Height), func, context);
            }

            // Both are aggregates
            {
                var isLeftArray = leftCollection.TryPickT0(out var leftArray, out var leftReference);
                var isRightArray = rightCollection.TryPickT0(out var rightArray, out var rightReference);

                if (isLeftArray && isRightArray)
                    return leftArray.Apply(rightArray, func, context);

                if (isLeftArray)
                {
                    if (rightReference.TryGetSingleCellValue(out var rightCellValue, context))
                        return leftArray.Apply(new ScalarArray(rightCellValue, leftArray.Width, leftArray.Height), func, context);

                    if (rightReference.Areas.Count == 1)
                        return leftArray.Apply(new ReferenceArray(rightReference.Areas[0], context), func, context);

                    return leftArray.Apply(new ScalarArray(XLError.IncompatibleValue, leftArray.Width, leftArray.Height), func, context);
                }

                if (isRightArray)
                {
                    if (leftReference.TryGetSingleCellValue(out var leftCellValue, context))
                        return new ScalarArray(leftCellValue, rightArray.Width, rightArray.Height).Apply(rightArray, func, context);

                    if (leftReference.Areas.Count == 1)
                        return new ReferenceArray(leftReference.Areas[0], context).Apply(rightArray, func, context);

                    return new ScalarArray(XLError.IncompatibleValue, rightArray.Width, rightArray.Height).Apply(rightArray, func, context);
                }

                // Both are references
                if (leftReference.Areas.Count > 1 && rightReference.Areas.Count > 1)
                    return XLError.IncompatibleValue;

                if (leftReference.Areas.Count > 1)
                    return new ScalarArray(XLError.IncompatibleValue, rightReference.Areas[0].ColumnSpan, rightReference.Areas[0].RowSpan);

                if (rightReference.Areas.Count > 1)
                    return new ScalarArray(XLError.IncompatibleValue, leftReference.Areas[0].ColumnSpan, leftReference.Areas[0].RowSpan);

                var leftArea = leftReference.Areas[0];
                var rightArea = rightReference.Areas[0];
                if (leftArea.IsSingleCell() && rightArea.IsSingleCell())
                {
                    var leftCellValue = context.GetCellValue(leftArea.Worksheet, leftArea.FirstAddress.RowNumber, leftArea.FirstAddress.ColumnNumber);
                    var rightCellValue = context.GetCellValue(rightArea.Worksheet, rightArea.FirstAddress.RowNumber, rightArea.FirstAddress.ColumnNumber);
                    return func(leftCellValue, rightCellValue, context).ToAnyValue();
                }

                var leftRefArray = new ReferenceArray(leftArea, context);
                var rightRefArray = new ReferenceArray(rightArea, context);
                return leftRefArray.Apply(rightRefArray, func, context);
            }
        }

        private static ScalarValue BinaryArithmeticOp(in ScalarValue left, in ScalarValue right, BinaryNumberFunc func, CalcContext ctx)
        {
            var leftConversionResult = left.ToNumber(ctx.Culture);
            if (!leftConversionResult.TryPickT0(out var leftNumber, out var leftError))
            {
                return leftError;
            }

            var rightConversionResult = right.ToNumber(ctx.Culture);
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
        private static OneOf<int, XLError> CompareValues(ScalarValue left, ScalarValue right, CultureInfo culture)
        {
            return left.Match(culture,
                _ => right.Match<OneOf<int, XLError>, CultureInfo>(culture,
                        _ => 0,
                        (rightLogical, _) => false.CompareTo(rightLogical),
                        (rightNumber, _) => 0.0.CompareTo(rightNumber),
                        (rightText, culture) => string.Compare(string.Empty, rightText, culture, CompareOptions.IgnoreCase),
                        (rightError, _) => rightError),
                (leftLogical, _) => right.Match<OneOf<int, XLError>, bool>(leftLogical,
                        leftLogical => leftLogical.CompareTo(false),
                        (rightLogical, leftLogical) => leftLogical.CompareTo(rightLogical),
                        (rightNumber, _) => 1,
                        (rightText, _) => 1,
                        (rightError, _) => rightError),
                (leftNumber, _) => right.Match<OneOf<int, XLError>, double>(leftNumber,
                        leftNumber => leftNumber.CompareTo(0.0),
                        (rightLogical, _) => -1,
                        (rightNumber, leftNumber) => leftNumber.CompareTo(rightNumber),
                        (rightText, _) => -1,
                        (rightError, _) => rightError),
                (leftText, culture) => right.Match<OneOf<int, XLError>, string, CultureInfo>(leftText, culture,
                        (leftText, culture) => string.Compare(leftText, string.Empty, culture, CompareOptions.IgnoreCase),
                        (rightLogical, _, _) => -1,
                        (rightNumber, _, _) => 1,
                        (rightText, leftText, culture) => string.Compare(leftText, rightText, culture, CompareOptions.IgnoreCase),
                        (rightError, _, _) => rightError),
                (leftError, _) => leftError);
        }

        private delegate OneOf<double, XLError> BinaryNumberFunc(double lhs, double rhs);
    }

    internal delegate ScalarValue BinaryFunc(in ScalarValue lhs, in ScalarValue rhs, CalcContext ctx);
}
