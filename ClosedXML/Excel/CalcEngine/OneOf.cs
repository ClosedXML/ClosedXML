using System;

namespace ClosedXML.Excel.CalcEngine
{
    internal readonly struct OneOf<T0, T1>
    {
        private readonly bool _isT0;
        private readonly T0 _t0;
        private readonly T1 _t1;

        private OneOf(bool isT0, T0 t0, T1 t1)
        {
            _isT0 = isT0;
            _t0 = t0;
            _t1 = t1;
        }

        public bool TryPickT0(out T0 t0, out T1 t1)
        {
            t0 = _t0;
            t1 = _t1;
            return _isT0;
        }

        public static OneOf<T0, T1> FromT0(T0 t0) => new(true, t0, default);

        public static OneOf<T0, T1> FromT1(T1 t1) => new(false, default, t1);

        public static implicit operator OneOf<T0, T1>(T0 t0) => FromT0(t0);

        public static implicit operator OneOf<T0, T1>(T1 t1) => FromT1(t1);

        public TResult Match<TResult>(Func<T0, TResult> transformT0, Func<T1, TResult> transformT1)
        {
            return _isT0 ? transformT0(_t0) : transformT1(_t1);
        }
    }
}
