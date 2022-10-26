// Keep this file CodeMaid organized and cleaned
using System;
using System.Drawing;
using ClosedXML.Graphics;

namespace ClosedXML.Excel
{
    /// <summary>
    /// A class that defines various aspects of a newly created workbook.
    /// </summary>
    public class LoadOptions
    {
        private Point _dpi = new(96, 96);

        /// <summary>
        /// A graphics engine that will be used for workbooks without explicitly set engine.
        /// </summary>
        public static IXLGraphicEngine DefaultGraphicEngine { internal get; set; }

        /// <summary>
        /// Should all formulas in a workbook be recalculated during load? Default value is <c>false</c>.
        /// </summary>
        public Boolean RecalculateAllFormulas { get; set; } = false;

        /// <summary>
        /// Graphic engine used by the workbook.
        /// </summary>
        public IXLGraphicEngine GraphicEngine { get; set; }

        /// <summary>
        /// DPI for the workbook. Default is 96.
        /// </summary>
        /// <remarks>Used in various places, e.g. determining a physical size of an image without a DPI or to determine a size of a text in a cell.</remarks>
        public Point Dpi
        {
            get => _dpi;
            set => _dpi = value.X > 0 && value.Y > 0 ? value : throw new ArgumentException("DPI must be positive");
        }
    }
}
