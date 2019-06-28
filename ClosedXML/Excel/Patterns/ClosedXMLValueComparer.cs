using ClosedXML.Excel.CalcEngine;
using System;
using System.Collections.Generic;

namespace ClosedXML.Excel.Patterns
{
    internal class ClosedXMLValueComparer : IEqualityComparer<object>, IComparer<object>
    {
        private static ClosedXMLValueComparer _defaultComparer = new ClosedXMLValueComparer(StringComparer.OrdinalIgnoreCase);
        private readonly StringComparer _stringComparer;

        private ClosedXMLValueComparer()
            : this(StringComparer.OrdinalIgnoreCase)
        { }

        private ClosedXMLValueComparer(StringComparer stringComparer)
        {
            this._stringComparer = stringComparer;
        }

        public static ClosedXMLValueComparer DefaultComparer { get { return _defaultComparer; } }

        public int Compare(object x, object y)
        {
            IComparable c1;
            if (x is Expression e1)
                c1 = e1.Evaluate() as IComparable;
            else
                c1 = x as IComparable;

            IComparable c2;
            if (y is Expression e2)
                c2 = e2.Evaluate() as IComparable;
            else
                c2 = y as IComparable;

            // handle nulls
            if (c1 == null && c2 == null)
                return 0;
            if (c2 == null)
                return -1;
            if (c1 == null)
                return +1;

            // make sure types are the same
            if (c1.GetType() != c2.GetType())
                return CompareWithCoersion(c1, c2);
            else
                return CompareWithoutCoersion(c1, c2);
        }

        public new bool Equals(object x, object y)
        {
            if (x == null && y == null)
                return true;

            if (y == null || x == null)
                return false;

            if (ReferenceEquals(x, y))
                return true;

            return Compare(x, y) == 0;
        }

        public int GetHashCode(object obj)
        {
            if (obj is null)
                return 0;

            switch (obj)
            {
                case DateTime dt:
                    return dt.ToOADate().GetHashCode();
                case Double dbl:
                    return dbl.GetHashCode();
                case TimeSpan ts:
                    return ts.TotalDays.GetHashCode();
                case Boolean b:
                    return (b ? 1d : 0d).GetHashCode();
                case String s:
                    return StringComparer.OrdinalIgnoreCase.GetHashCode(s);
                default:
                    throw new NotImplementedException();
            }
        }

        private int CompareWithCoersion(IComparable c1, IComparable c2)
        {
            try
            {
                if (c2 is DateTime dt2 && c1.IsNumber())
                    c1 = Convert.ChangeType(new ConvertibleObject(c1), typeof(DateTime)) as IComparable;
                else
                    c2 = Convert.ChangeType(new ConvertibleObject(c2), c1.GetType()) as IComparable;

                return CompareWithoutCoersion(c1, c2);
            }
            catch (InvalidCastException) { return -1; }
            catch (FormatException) { return -1; }
            catch (OverflowException) { return -1; }
            catch (ArgumentNullException) { return -1; }
        }

        private int CompareWithoutCoersion<T>(T c1, T c2)
            where T : IComparable
        {
            if (c1 is string s1 && c2 is string s2)
                return _stringComparer.Compare(s1, s2);

            return c1.CompareTo(c2);
        }
    }
}
