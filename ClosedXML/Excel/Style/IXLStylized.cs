using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    /// <summary>
    /// An interface implemented by workbook elements that have a defined <see cref="IXLStyle"/>.
    /// </summary>
    internal interface IXLStylized
    {
        /// <summary>
        /// Editable style of the workbook element. Modification of this property DOES affect styles of child objects as well - they will
        /// be changed accordingly. Accessing this property causes a new <see cref="XLStyle"/> instance generated so use this property
        /// with caution. If you need only _read_ the style consider using <see cref="StyleValue"/> property instead.
        /// </summary>
        IXLStyle Style { get; set; }

        /// <summary>
        /// Editable style of the workbook element. Modification of this property DOES NOT affect styles of child objects.
        /// Accessing this property causes a new <see cref="XLStyle"/> instance generated so use this property with caution. If you need
        /// only _read_ the style consider using <see cref="StyleValue"/> property instead.
        /// </summary>
        IXLStyle InnerStyle { get; set; }

        /// <summary>
        /// <para>
        /// Return a collection of ranges that determine outside borders (used by
        /// <see cref="XLBorder.OutsideBorder"/>).
        /// </para>
        /// <para>
        /// Return ranges represented by elements. For one element (e.g. workbook, cell,
        /// column), it should return only the element itself. For element that represent a
        /// collection of other elements, e.g. <see cref="XLRows"/>, <see cref="XLColumns"/>,
        /// <see cref="XLCells"/>, it should return range for each element in the collection.
        /// </para>
        /// </summary>
        IXLRanges RangesUsed { get; }

        /// <summary>
        /// Style value representing the current style of the stylized element.
        /// The value is updated when style is modified (<see cref="XLStyleValue"/>
        /// is immutable).
        /// </summary>
        XLStyleValue StyleValue { get; }

        /// <summary>
        /// A callback method called when <see cref="Style"/> is changed. It should update
        /// style of the stylized descendants of the stylized element.
        /// </summary>
        /// <param name="modification">A method that changes the style from original to modified.</param>
        void ModifyStyle(Func<XLStyleKey, XLStyleKey> modification);
    }
}
