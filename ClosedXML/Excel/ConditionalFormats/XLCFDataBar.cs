#nullable disable

using System;
using System.Collections.Generic;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLCFDataBar : IXLCFDataBar
    {
        private readonly XLConditionalFormat _conditionalFormat;

        public XLCFDataBar(XLConditionalFormat conditionalFormat)
        {
            _conditionalFormat = conditionalFormat;
            _conditionalFormat.DataBar = this;
            // Default value in Excel is true
            Gradient = true;
        }

        /// <summary>
        /// The fill type of a DataBar, Gradient (true) or Solid.
        /// </summary>
        public bool Gradient { get; set; }

        /// <summary>
        /// Specifies the fill type of a DataBar, Gradient or Solid.
        /// </summary>
        /// <param name="value">true to apply a gradient fill (default), false to apply a solid fill</param>
        /// <returns>The <see cref="IXLCFDataBar" /></returns>
        public IXLCFDataBar SetGradient(bool value = true)
        {
            Gradient = value;
            return this;
        }
    }
}
