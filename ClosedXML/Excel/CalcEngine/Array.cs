using System;
using System.Collections.Generic;
using System.Linq;
using ScalarValue = OneOf.OneOf<bool, double, string, ClosedXML.Excel.CalcEngine.Error>;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// A base class for an 2D array. Every array is at least 1x1.
    /// </summary>
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

        /// <summary>
        /// An iterator over all elements of an array, from top to bottom, from left to right.
        /// </summary>
        public virtual IEnumerator<ScalarValue> GetEnumerator() => FlattenArray().GetEnumerator();

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();

        protected IEnumerable<ScalarValue> FlattenArray()
        {
            for (int row = 0; row < Height; row++)
            {
                for (int col = 0; col < Width; col++)
                {
                    yield return this[row, col];
                }
            }
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
                if (x < 0 || x >= _width || y < 0 || y >= _height)
                    throw new IndexOutOfRangeException();

                return _value;
            }
        }

        public override IEnumerator<ScalarValue> GetEnumerator()
        {
            return Enumerable.Range(0, _width * _height).Select(_ => _value).GetEnumerator();
        }
    }

    /// <summary>
    /// An array of scalar values.
    /// </summary>
    internal class ConstArray : Array
    {
        internal readonly ScalarValue[,] _data;

        public ConstArray(ScalarValue[,] data)
        {
            if (data.GetLength(0) < 1 || data.GetLength(1) < 1)
                throw new ArgumentException("Array must be at least 1x1.", nameof(data));
            _data = data;
        }

        public override ScalarValue this[int y, int x] => _data[y, x];

        public override int Width => _data.GetLength(1);

        public override int Height => _data.GetLength(0);
    }

    /// <summary>
    /// A special case of an array that is actually only numbers.
    /// </summary>
    internal class NumberArray : Array
    {
        private readonly double[,] _data;

        public NumberArray(double[,] data)
        {
            _data = data;
        }

        public override ScalarValue this[int y, int x] => _data[y, x];

        public override int Width => _data.GetLength(1);

        public override int Height => _data.GetLength(0);
    }

    /// <summary>
    /// An array that is resized to a different size. Items outside of original array have a value of <c>#N/A</c>.
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
                    return Error.NoValueAvailable;

                return _original[y, x];
            }
        }
    }

    /// <summary>
    /// An array that retrieves its value directly from the worksheet without allocating extra memory.
    /// </summary>
    internal class ReferenceArray : Array
    {
        private readonly XLRangeAddress _area;
        private readonly CalcContext _context;
        private readonly int _offsetColumn;
        private readonly int _offsetRow;

        public ReferenceArray(XLRangeAddress area, CalcContext context)
        {
            _area = area;
            _context = context;
            _offsetColumn = _area.FirstAddress.ColumnNumber;
            _offsetRow = area.FirstAddress.RowNumber;
        }

        public override ScalarValue this[int y, int x] => _context.GetCellValue(_area.Worksheet, y + _offsetRow, x + _offsetColumn);

        public override int Width => _area.ColumnSpan;

        public override int Height => _area.RowSpan;
    }
}
