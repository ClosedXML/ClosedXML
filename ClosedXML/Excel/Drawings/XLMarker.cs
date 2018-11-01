// Keep this file CodeMaid organised and cleaned
using System;
using System.Diagnostics;

namespace ClosedXML.Excel.Drawings
{
    [DebuggerDisplay("{Address} {Offset}")]
    internal class XLMarker
    {
        // Using a range to store the location so that it gets added to the range repository
        // and hence will be adjusted when there are insertions / deletions
        private readonly IXLRange rangeCell;

        internal XLMarker(IXLCell cell)
            : this(cell.AsRange(), XLMeasure.Zero, XLMeasure.Zero)
        { }

        internal XLMarker(IXLCell cell, IXLMeasure xOffset, IXLMeasure yOffset)
            : this(cell.AsRange(), xOffset, yOffset)
        { }

        private XLMarker(IXLRange rangeCell, IXLMeasure xOffset, IXLMeasure yOffset)
        {
            if (rangeCell.RowCount() != 1 || rangeCell.ColumnCount() != 1)
                throw new ArgumentException("Range should contain only one cell.", nameof(rangeCell));

            this.rangeCell = rangeCell;
            this.X = xOffset;
            this.Y = yOffset;
        }

        public IXLCell Cell { get => rangeCell.FirstCell(); }
        public Int32 ColumnNumber { get => rangeCell.RangeAddress.FirstAddress.ColumnNumber; }
        public Int32 RowNumber { get => rangeCell.RangeAddress.FirstAddress.RowNumber; }
        public IXLMeasure X { get; set; }
        public IXLMeasure Y { get; set; }
    }
}
