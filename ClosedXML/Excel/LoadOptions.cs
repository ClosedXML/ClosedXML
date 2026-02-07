// Keep this file CodeMaid organized and cleaned
using System;
using System.Drawing;
using System.Threading;
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
        public static IXLGraphicEngine? DefaultGraphicEngine { internal get; set; }

        /// <summary>
        /// Should all formulas in a workbook be recalculated during load? Default value is <c>false</c>.
        /// </summary>
        public Boolean RecalculateAllFormulas { get; set; } = false;

        /// <summary>
        /// Graphic engine used by the workbook.
        /// </summary>
        public IXLGraphicEngine? GraphicEngine { get; set; }

        /// <summary>
        /// DPI for the workbook. Default is 96.
        /// </summary>
        /// <remarks>Used in various places, e.g. determining a physical size of an image without a DPI or to determine a size of a text in a cell.</remarks>
        public Point Dpi
        {
            get => _dpi;
            set => _dpi = value.X > 0 && value.Y > 0 ? value : throw new ArgumentException("DPI must be positive");
        }

        /// <summary>
        /// <para>
        /// The option determines how should the XML parser handle XML attribute whose values do
        /// not match the expected XML simple type (e.g., attribute value is <c>one</c>, but it
        /// should be a <c>xsd:integer</c> or attribute value should be an enum <c>ST_BorderStyle</c>,
        /// but contains invalid value  <c>triangle</c> instead). If the option is enabled (default),
        /// the XML parser will throw when in encounters an attribute with a value that violates its
        /// simple type definition. <em><b>Disabling the option will hide these errors</b></em>.
        /// </para>
        /// <para>
        /// <list type="bullet">
        ///   <item>
        ///     <c>true</c> - the XML parser will throw when it encounters attribute with value
        ///     that doesn't match its simple type.
        ///   </item>
        ///   <item>
        ///     <c>false</c> - the XML parser will interpret attribute with invalid value as if
        ///     the attribute wasn't present.
        ///   </item>
        /// </list>
        /// Disabling this option can help in following cases: attribute is optional (malformed
        /// value is interpreted as missing) or attribute is optional, but has a default value
        /// (malformed value is interpreted as missing and a default value is used instead).
        /// </para>
        /// </summary>
        /// <remarks>
        /// This option only affects <c>ClosedXML.IO</c> parser. OOXML parts that are parsed by
        /// OpenXML SDK are not affected. Also, fix an OOXML producer that does this stuff.
        /// </remarks>
        public bool StrictAttributeParsing { get; set; } = true;

        /// <summary>
        /// A cancellation token to cancel recalculation. Recalculation may be triggered during
        /// various calls, e.g. <c>IXLCell.Value</c> can start recalculation of a workbook, but
        /// doesn't have a way to pass token.
        /// </summary>
        public CancellationToken CancellationToken { get; set; }
    }
}
