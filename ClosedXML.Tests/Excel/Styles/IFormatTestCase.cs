using System.Collections.Generic;

namespace ClosedXML.Tests.Excel.Styles;

/// <summary>
/// A test case for setting a format property.
/// </summary>
/// <typeparam name="T">Type of property that is being set.</typeparam>
public interface IFormatTestCase<in T>
{
    /// <summary>
    /// Test values.
    /// </summary>
    IEnumerable<object> Values { get; }

    object GetPropertyValue(T font);

    void SetPropertyValue(T font, object value);
}
