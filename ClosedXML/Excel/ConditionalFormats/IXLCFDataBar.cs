#nullable disable

using System;
using System.Collections.Generic;
using System.Text;

namespace ClosedXML.Excel
{
    /// <summary>
    /// A fluent configuration builder for the DataBar object
    /// </summary>
    public interface IXLCFDataBar
    {
        /// <summary>
        /// The fill type of a DataBar, Gradient (true) or Solid.
        /// </summary>
        bool Gradient { get; }

        /// <summary>
        /// Specifies the fill type of a DataBar, Gradient or Solid.
        /// </summary>
        /// <param name="value">true to apply a gradient fill (default), false to apply a solid fill</param>
        /// <returns>The <see cref="IXLCFDataBar" /></returns>
        IXLCFDataBar SetGradient(bool value = true);
    }
}
