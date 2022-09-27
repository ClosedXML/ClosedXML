// Keep this file CodeMaid organized and cleaned
using System;
using System.Drawing;

namespace ClosedXML.Excel
{
    /// <summary>
    /// A class that defines various aspects of a newly created workbook.
    /// </summary>
    public class LoadOptions
    {
        private Point _dpi = new(96, 96);

        public XLEventTracking EventTracking { get; set; } = XLEventTracking.Enabled;

        /// <summary>
        /// Should all formulas in a workbook be recalculated during load? Default value is <c>false</c>.
        /// </summary>
        public Boolean RecalculateAllFormulas { get; set; } = false;

        /// <summary>
        /// DPI for the workbook. Default is 96.
        /// </summary>
        /// <remarks>Used in various places, e.g. determining a physical size of an image without a DPI or to determine a size of a text in a cell.</remarks>
        public Point Dpi
        {
            get => _dpi;
            set => _dpi = _dpi.X > 0 && _dpi.Y > 0 ? value : throw new ArgumentException("DPI must be positive");
        }
    }
}
