using System;
using System.Threading;

namespace ClosedXML.Excel.Patterns
{
    internal class ConvertibleObject : IComparable, IConvertible
    {
        public ConvertibleObject(Object value)
        {
            Value = value;
        }

        public Object Value { get; }

        public int CompareTo(object obj)
        {
            switch (this.Value)
            {
                case Double dbl:
                    return dbl.CompareTo((Double)obj);

                case DateTime dt:
                    return dt.CompareTo((DateTime)obj);

                case Boolean b:
                    return b.CompareTo((Boolean)obj);

                case TimeSpan ts:
                    return ts.CompareTo((TimeSpan)obj);

                case String s:
                    return StringComparer.OrdinalIgnoreCase.Compare(s, (String)obj);

                default:
                    throw new NotImplementedException();
            }
        }

        #region IConvertible interface

        public TypeCode GetTypeCode()
        {
            throw new NotImplementedException();
        }

        public bool ToBoolean(IFormatProvider provider)
        {
            return (bool)this;
        }

        public byte ToByte(IFormatProvider provider)
        {
            throw new NotImplementedException();
        }

        public char ToChar(IFormatProvider provider)
        {
            throw new NotImplementedException();
        }

        public DateTime ToDateTime(IFormatProvider provider)
        {
            return (DateTime)this;
        }

        public decimal ToDecimal(IFormatProvider provider)
        {
            throw new NotImplementedException();
        }

        public double ToDouble(IFormatProvider provider)
        {
            return (Double)this;
        }

        public short ToInt16(IFormatProvider provider)
        {
            return Convert.ToInt16((Double)this);
        }

        public int ToInt32(IFormatProvider provider)
        {
            return Convert.ToInt32((Double)this);
        }

        public long ToInt64(IFormatProvider provider)
        {
            return Convert.ToInt64((Double)this);
        }

        public sbyte ToSByte(IFormatProvider provider)
        {
            throw new NotImplementedException();
        }

        public float ToSingle(IFormatProvider provider)
        {
            throw new NotImplementedException();
        }

        public string ToString(IFormatProvider provider)
        {
            return (String)this;
        }

        public object ToType(Type conversionType, IFormatProvider provider)
        {
            throw new NotImplementedException();
        }

        public ushort ToUInt16(IFormatProvider provider)
        {
            throw new NotImplementedException();
        }

        public uint ToUInt32(IFormatProvider provider)
        {
            throw new NotImplementedException();
        }

        public ulong ToUInt64(IFormatProvider provider)
        {
            throw new NotImplementedException();
        }

        #endregion IConvertible interface

        #region ** implicit converters

        public static implicit operator bool(ConvertibleObject x)
        {
            // handle nulls
            if (x.Value == null)
                return false;

            // handle booleans
            if (x.Value is bool b)
                return b;

            // handle doubles
            if (x.Value is double dbl)
                return Math.Abs(dbl) > double.Epsilon;

            // handle everything else
            return (double)Convert.ChangeType(x.Value, typeof(double)) != 0;
        }

        public static implicit operator DateTime(ConvertibleObject x)
        {
            // handle dates
            if (x.Value is DateTime dt)
                return dt;

            if (x.Value is TimeSpan ts)
                return new DateTime().Add(ts);

            // handle numbers
            if (x.Value.IsNumber())
                return DateTime.FromOADate(Convert.ToDouble(x.Value));

            // handle everything else
            var _ci = Thread.CurrentThread.CurrentCulture;
            return (DateTime)Convert.ChangeType(x.Value, typeof(DateTime), _ci);
        }

        public static implicit operator double(ConvertibleObject x)
        {
            // handle doubles
            if (x.Value is double dbl)
                return dbl;

            // handle booleans
            if (x.Value is bool b)
                return b ? 1 : 0;

            // handle dates
            if (x.Value is DateTime dt)
                return dt.ToOADate();

            if (x.Value is TimeSpan ts)
                return ts.TotalDays;

            // handle nulls
            if (x.Value == null || x.Value is string)
                return 0;

            // handle everything else
            var _ci = Thread.CurrentThread.CurrentCulture;
            return (double)Convert.ChangeType(x.Value, typeof(double), _ci);
        }

        public static implicit operator string(ConvertibleObject x)
        {
            if (x.Value == null)
                return string.Empty;

            if (x.Value is bool b)
                return b.ToString().ToUpper();

            return x.Value.ToInvariantString();
        }

        #endregion ** implicit converters
    }
}
