using OneOf;
using System;
using System.Collections.Generic;
using System.Linq;
using AnyValue = OneOf.OneOf<ClosedXML.Excel.CalcEngine.Logical, ClosedXML.Excel.CalcEngine.Number1, ClosedXML.Excel.CalcEngine.Text, ClosedXML.Excel.CalcEngine.Error1, ClosedXML.Excel.CalcEngine.Array, ClosedXML.Excel.CalcEngine.Reference>;
using ScalarValue = OneOf.OneOf<ClosedXML.Excel.CalcEngine.Logical, ClosedXML.Excel.CalcEngine.Number1, ClosedXML.Excel.CalcEngine.Text, ClosedXML.Excel.CalcEngine.Error1>;

namespace ClosedXML.Excel.CalcEngine
{
    internal readonly struct Logical
    {
        public static readonly Logical True = new(true);
        public static readonly Logical False = new(false);

        public Logical(bool value) => Value = value;

        public bool Value { get; }

        public static implicit operator bool(Logical logical) => logical.Value;
        public static implicit operator Logical(bool value) => new(value);

        public override string ToString() => Value.ToString();

        public static Logical And(Logical lhs, Logical rhs)
        {
            throw new NotImplementedException();
        }
    }

    internal readonly struct Number1
    {
        public static readonly Number1 Zero = new(0.0);
        public static readonly Number1 One = new(1.0);

        public Number1(double value) => Value = value;

        public double Value { get; }

        public static implicit operator double(Number1 number) => number.Value;
        public static implicit operator Number1(double value) => new Number1(value);

        public override string ToString() => Value.ToString();

        public static Number1 Plus(Number1 lhs, Number1 rhs)
        {
            throw new NotImplementedException();
        }
    }

    internal readonly struct Text
    {
        public Text(string value)
        {
            if (value is null) throw new ArgumentNullException(nameof(value));
            Value = value;
        }

        public string Value { get; }

        public static implicit operator string(Text text) => text.Value;
        public static implicit operator Text(string value) => new(value);

        public override string ToString() => Value.ToString();

        public static Text Concat(Text lhs, Text rhs)
        {
            throw new NotImplementedException();
        }
    }

    // There is no downside to have custom type, we can add more info, if we want (like text)
    internal readonly struct Error1
    {
        /// <summary>
        /// #VALUE!
        /// </summary>
        /// <remarks>Intended to indicate when an incompatible type argument is passed to a function, or an incompatible type operand is used with an operator.</remarks>
        public static readonly Error1 Value = new(ExpressionErrorType.CellValue);

        /// <summary>
        /// #DIV/0!
        /// </summary>
        public static readonly Error1 DivZero = new(ExpressionErrorType.DivisionByZero);

        /// <summary>
        /// #NUM!
        /// </summary>
        public static readonly Error1 NumberInvalid = new(ExpressionErrorType.NumberInvalid);

        /// <summary>
        /// #N/A
        /// </summary>
        public static readonly Error1 NoValueAvailable = new(ExpressionErrorType.NoValueAvailable);

        /// <summary>
        /// #REF!
        /// </summary>
        public static readonly Error1 Ref = new(ExpressionErrorType.CellReference);

        /// <summary>
        /// #NAME?
        /// </summary>
        public static readonly Error1 Name = new(ExpressionErrorType.NameNotRecognized);

        /// <summary>
        /// #NULL!
        /// </summary>
        public static readonly Error1 Null = new(ExpressionErrorType.NullValue);

        public Error1(ExpressionErrorType type) => Type = type;

        public ExpressionErrorType Type { get; }

        public override string ToString() => Type.ToString();
    }

    // 2D array of values, always at least 1x1
    internal abstract class Array : IEnumerable<ScalarValue>
    {
        /// <summary>
        /// Width of the array, at least 1.
        /// </summary>
        public abstract int Width { get; }

        /// <summary>
        /// Height of the array, at least 1.
        /// </summary>
        public abstract int Height { get; }

        /// <summary>
        /// Get a value at specified coordinate.
        /// </summary>
        /// <param name="y">Uses 0-based notation.</param>
        /// <param name="x">Uses 0-based notation.</param>
        public abstract ScalarValue this[int y, int x] { get; }

        // get iterator over all elements of an array, from top to bottom, from left to right.
        public abstract IEnumerator<ScalarValue> GetEnumerator();

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();

        public static Array Plus(Array lhs, Array rhs)
        {
            throw new NotImplementedException();
        }
    }

    /// <summary>
    /// An array that is using a different array to store data. This way, we can keep memory pressures at minimum.
    /// This is intermediate array used during calculations.
    /// </summary>
    internal class CalculatedArray : Array
    {
        private readonly Array _array;
        private readonly Func<ScalarValue, ScalarValue> _func;

        public CalculatedArray(Array array, Func<ScalarValue, ScalarValue> func)
        {
            _array = array;
            _func = func;
        }

        public override int Width => _array.Width;

        public override int Height => _array.Height;

        public override ScalarValue this[int y, int x] => _func(_array[y, x]);

        public override IEnumerator<ScalarValue> GetEnumerator() => _array.GetEnumerator();
    }

    internal class ScalarArray : Array
    {
        private readonly ScalarValue _value;
        private readonly int _width;
        private readonly int _height;

        public ScalarArray(ScalarValue value, int width, int height)
        {
            if (width < 1) throw new ArgumentOutOfRangeException(nameof(width));
            if (height < 1) throw new ArgumentOutOfRangeException(nameof(height));
            _value = value;
            _width = width;
            _height = height;
        }

        public override int Width => _width;

        public override int Height => _height;

        public override ScalarValue this[int y, int x]
        {
            get
            {
                if (x < 0 || x >= _width) throw new IndexOutOfRangeException();
                if (y < 0 || y >= _height) throw new IndexOutOfRangeException();
                return _value;
            }
        }

        public override IEnumerator<ScalarValue> GetEnumerator()
        {
            throw new NotImplementedException();
        }
    }

    internal class ConstArray : Array
    {
        internal readonly ScalarValue[,] _data;

        public ConstArray(ScalarValue[,] data)
        {
            if (data.GetLength(0) < 1 || data.GetLength(1) < 1)
                throw new ArgumentException("Array must be at least 1x1.", nameof(data));
            _data = data;
        }

        public override ScalarValue this[int x, int y] => _data[x, y];

        public override int Width => _data.GetLength(1);

        public override int Height => _data.GetLength(0);

        public override IEnumerator<ScalarValue> GetEnumerator() => _data.Cast<ScalarValue>().GetEnumerator();
    }

    /// <summary>
    /// An array that is enlarged
    /// </summary>
    internal class ResizedArray : Array
    {
        private readonly Array _original;

        public ResizedArray(Array original, int width, int height)
        {
            if (width < 1 || height < 1)
                throw new ArgumentException();

            _original = original;
            Width = width;
            Height = height;
        }

        public override int Width { get; }

        public override int Height { get; }

        public override ScalarValue this[int y, int x]
        {
            get
            {
                if (x >= _original.Width || y >= _original.Height)
                    return ScalarValue.FromT3(Error1.NoValueAvailable);

                return _original[y, x];
            }
        }
        public override IEnumerator<ScalarValue> GetEnumerator()
        {
            throw new NotImplementedException();
        }
    }

    internal class ReferenceArray : Array
    {
        private readonly XLRangeAddress _area;
        private readonly CalcContext _context;

        public ReferenceArray(XLRangeAddress area, CalcContext context)
        {
            _area = area;
            _context = context;
        }

        public override ScalarValue this[int y, int x]
        {
            get
            {
                return AnyValueExtensions.GetCellValue(_area, y + 1, x + 1, _context);
            }
        }

        public override int Width => _area.ColumnSpan;

        public override int Height => _area.RowSpan;

        public override IEnumerator<ScalarValue> GetEnumerator()
        {
            throw new NotImplementedException();
        }
    }

    // 2D area on a worksheet that is a part of Reference
    internal class Area
    {
        /// <summary>Columna and row are global</summary>
        public Area(XLWorksheet worksheet, int column, int row, int width, int height)
        {
            if (worksheet is null)
                throw new ArgumentNullException(nameof(worksheet));
            if (column < 1 || column > XLConstants.MaxColumns)
                throw new ArgumentOutOfRangeException(nameof(column));
            if (row < 1 || row > XLConstants.MaxRows)
                throw new ArgumentOutOfRangeException(nameof(row));

            Worksheet = worksheet;
            Column = column;
            Row = row;
            Width = width;
            Height = height;
        }



        public XLWorksheet Worksheet { get; }

        /// <summary>
        /// 1 based
        /// </summary>
        public int Column { get; }

        public int Row { get; }

        public int Width { get; }

        public int Height { get; }
    }
}
