using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace ClosedXML.Tests.Excel.Styles;

public class FormatTestCase<TApi>
{
    private readonly Func<TApi, object> _getter;
    private readonly Action<TApi, object> _setter;
    private readonly IReadOnlyList<object> _testValues;

    private FormatTestCase(Func<TApi, object> getter, Action<TApi, object> setter, params object[] testValues)
    {
        _getter = getter;
        _setter = setter;
        _testValues = testValues;
    }

    public static FormatTestCase<IXLFont> ForFont<T>(Func<IXLFont, T> getter, Action<IXLFont, T> setter, params T[] testValues)
    {
        return new FormatTestCase<IXLFont>(font => getter(font), (font, value) => setter(font, (T)value), testValues.Cast<object>().ToArray());
    }

    public static FormatTestCase<IXLFill> ForFill<T>(Func<IXLFill, T> getter, Action<IXLFill, T> setter, params T[] testValues)
    {
        return new FormatTestCase<IXLFill>(fill => getter(fill), (fill, value) => setter(fill, (T)value), testValues.Cast<object>().ToArray());
    }

    public IEnumerable<object> Values => _testValues;

    public object GetPropertyValue(TApi font)
    {
        return _getter(font);
    }

    public void SetPropertyValue(TApi font, object value)
    {
        _setter(font, value);
    }
}
