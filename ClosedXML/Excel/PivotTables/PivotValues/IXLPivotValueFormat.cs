using System;

namespace ClosedXML.Excel
{
    /// <summary>
    /// An API for manipulating a <see cref="IXLPivotValue.NumberFormat">format</see> of one
    /// <see cref="IXLPivotTable"/> <see cref="IXLPivotValue">data field</see>.
    /// </summary>
    public interface IXLPivotValueFormat : IXLNumberFormatBase
    {
        /// <summary>
        /// Set number formatting using one of predefined codes. Predefined codes are described in
        /// the <see cref="XLPredefinedFormat"/>.
        /// </summary>
        /// <param name="value">A numeric value describing how should the number be formatted.</param>
        IXLPivotValue SetNumberFormatId(Int32 value);

        IXLPivotValue SetFormat(String value);
    }
}
