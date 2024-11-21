#nullable disable

using System;
using System.Globalization;
using System.Linq;
using ClosedXML.Extensions;
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

        public bool IsLogical => _index == LogicalValue;

        public bool IsNumber => _index == NumberValue;

        public bool IsText => _index == TextValue;

        public bool IsError => _index == ErrorValue;

        public bool IsArray => _index == ArrayValue;

        public bool IsReference => _index == ReferenceValue;

        /// <summary>
        /// Is the value a scalar (blank, logical, number, text or error).
        /// </summary>
        public bool IsScalarType => IsBlank || IsLogical || IsNumber || IsText || IsError;

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

        public bool TryPickArray(out Array array)
        {
            if (_index == ArrayValue)
            {
                array = _array;
                return true;
            }

            array = default;
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

        /// <summary>
        /// Return array from a single area reference or array. If value is scalar, return false.
        /// </summary>
        public bool TryPickCollectionArray(out Array array, CalcContext ctx)
        {
            if (TryPickArea(out var areaAddress, out _))
            {
                array = new ReferenceArray(areaAddress, ctx);
                return true;
            }

            if (TryPickArray(out array))
                return true;

            array = null;
            return false;
        }

        /// <summary>
        /// <para>
        /// Try to get a value more in line with an array formula semantic. The output is always
        /// either single value or an array.
        /// </para>
        /// <para>
        /// Single cell references are turned into a scalar, multi-area references are turned
        /// into <see cref="XLError.IncompatibleValue"/> and single-area references are turned
        /// into arrays.
        /// </para>
        /// </summary>
        /// <remarks>
        /// Note the difference in nomenclature: <em>single/multi value</em> vs <em>scalar/collection type</em>.
        /// </remarks>
        internal bool TryPickSingleOrMultiValue(out ScalarValue scalar, out Array array, CalcContext ctx)
        {
            if (TryPickScalar(out scalar, out var collection))
            {
                array = default;
                return true;
            }

            // For some weird reason, 1x1 array doesn't count as a scalar, unlike single cell reference
            // proof {=TYPE(A1+1)} is 1 (scalar), but {=TYPE({1}+1)} is 64 (array).
            if (collection.TryPickT0(out array, out var reference))
            {
                scalar = default;
                return false;
            }

            if (reference.TryGetSingleCellValue(out scalar, ctx))
            {
                return true;
            }

            if (reference.Areas.Count > 1)
            {
                scalar = XLError.IncompatibleValue;
                return true;
            }

            array = new ReferenceArray(reference.Areas[0], ctx);
            return false;
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
            var isSingle = value.TryPickSingleOrMultiValue(out var single, out var array, context);
            if (isSingle)
                return UnaryArithmeticOp(single, operatorFn, context).ToAnyValue();

            return array.Apply(arrayConst => UnaryArithmeticOp(arrayConst, operatorFn, context));
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
            var isLeftSingle = left.TryPickSingleOrMultiValue(out var leftSingle, out var leftArray, context);
            var isRightSingle = right.TryPickSingleOrMultiValue(out var rightSingle, out var rightArray, context);

            if (isLeftSingle && isRightSingle)
                return func(in leftSingle, in rightSingle, context).ToAnyValue();

            if (isLeftSingle)
            {
                var broadcastedLeftArray = new ScalarArray(leftSingle, rightArray.Width, rightArray.Height);
                return broadcastedLeftArray.Apply(rightArray, func, context);
            }

            if (isRightSingle)
            {
                var broadcastedRightArray = new ScalarArray(rightSingle, leftArray.Width, leftArray.Height);
                return leftArray.Apply(broadcastedRightArray, func, context);
            }

            var unifiedRows = Math.Max(leftArray.Height, rightArray.Height);
            var unifiedColumns = Math.Max(leftArray.Width, rightArray.Width);

            var leftBroadcastedArray = leftArray.Broadcast(unifiedRows, unifiedColumns);
            var rightBroadcastedArray = rightArray.Broadcast(unifiedRows, unifiedColumns);

            return leftBroadcastedArray.Apply(rightBroadcastedArray, func, context);
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

        public override string ToString()
        {
            return _index switch
            {
                BlankValue => "Blank",
                LogicalValue => $"Logical: {_logical.ToString().ToUpper()}",
                NumberValue => $"Number: {_number}",
                TextValue => $"Text: {_text}",
                ErrorValue => $"Error: {_error.ToDisplayString()}",
                ArrayValue => $"Array{_array.Height}x{_array.Width}",
                ReferenceValue => $"Reference: {string.Join(",", _reference.Areas.Select(a => $"{a.FirstAddress}:{a.LastAddress}"))}",
                _ => throw new InvalidOperationException()
            };
        }

        /// <summary>
        /// Get 2d size of the value. For scalars, it's 1x1, for multi-area references,
        /// it's also 1x1,because it is converted to <c>#VALUE!</c> error.
        /// </summary>
        public (int Rows, int Columns) GetArraySize()
        {
            if (IsScalarType)
                return (1, 1);

            if (TryPickArray(out var array))
                return (array.Height, array.Width);

            if (TryPickArea(out var area, out _))
                return (area.RowSpan, area.ColumnSpan);

            // Multi area is just error = scalar
            return (1, 1);
        }

        /// <summary>
        /// Return the array value.
        /// </summary>
        /// <exception cref="InvalidCastException" />
        public Array GetArray() => _index == ArrayValue ? _array : throw new InvalidCastException();

        private delegate OneOf<double, XLError> BinaryNumberFunc(double lhs, double rhs);
    }

    internal delegate ScalarValue BinaryFunc(in ScalarValue lhs, in ScalarValue rhs, CalcContext ctx);
}
