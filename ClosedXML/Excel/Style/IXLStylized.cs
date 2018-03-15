using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal interface IXLStylized
    {
        IXLStyle Style { get; set; }

        IEnumerable<IXLStyle> Styles { get; }

        IXLStyle InnerStyle { get; set; }

        IXLRanges RangesUsed { get; }

        /// <summary>
        /// Immutable style
        /// </summary>
        XLStyleValue StyleValue { get; }

        void ModifyStyle(Func<XLStyleKey, XLStyleKey> modification);
    }
}
