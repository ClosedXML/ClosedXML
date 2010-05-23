using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel.Style
{
    public enum XLAlignmentReadingOrderValues
    {
        ContextDependent,
        LeftToRight,
        RightToLeft
    }

    public enum XLAlignmentHorizontalValues
    {
        Center,
        CenterContinuous,
        Distributed,
        Fill,
        General,
        Justify,
        Left,
        Right
    }

    public enum XLAlignmentVerticalValues
    {
        Bottom,
        Center,
        Distributed,
        Justify,
        Top
    }

    public class XLAlignment
    {
        #region Properties

        private XLRange range;

        private XLAlignmentHorizontalValues horizontal;
        public XLAlignmentHorizontalValues Horizontal
        {
            get { return horizontal; }
            set
            {
                Boolean updateIndent = !(
                    value == XLAlignmentHorizontalValues.Left
                    || value == XLAlignmentHorizontalValues.Right
                    || value == XLAlignmentHorizontalValues.Distributed
                    );

                if (updateIndent)
                    indent = 0;

                horizontal = value;

                if (range != null) range.ProcessCells(c =>
                {
                    if (updateIndent)
                        c.CellStyle.Alignment.indent = 0;

                    c.CellStyle.Alignment.horizontal = value;
                });
            }
        }

        private XLAlignmentVerticalValues vertical;
        public XLAlignmentVerticalValues Vertical
        {
            get { return vertical; }
            set
            {
                vertical = value;
                if (range != null) range.ProcessCells(c => c.CellStyle.Alignment.vertical = value);
            }
        }

        private UInt32 indent;
        public UInt32 Indent
        {
            get
            {
                return indent;
            }
            set
            {
                if (value > 0 && !(
                    Horizontal == XLAlignmentHorizontalValues.Left
                    || Horizontal == XLAlignmentHorizontalValues.Right
                    || Horizontal == XLAlignmentHorizontalValues.Distributed
                    ))
                {
                    throw new ArgumentException("For indents, only left, right, and distributed horizontal alignments are supported.");
                }
                indent = value;

                if (range != null) range.ProcessCells(c =>
                {
                    if (value > 0 && !(
                        c.CellStyle.Alignment.horizontal == XLAlignmentHorizontalValues.Left
                        || c.CellStyle.Alignment.horizontal == XLAlignmentHorizontalValues.Right
                        || c.CellStyle.Alignment.horizontal == XLAlignmentHorizontalValues.Distributed
                        ))
                    {
                        throw new ArgumentException("For indents, only left, right, and distributed horizontal alignments are supported. Change the horizontal alignment for all cells in the range.");
                    }

                    c.CellStyle.Alignment.indent = value;
                });
            }
        }

        private Boolean justifyLastLine;
        public Boolean JustifyLastLine
        {
            get { return justifyLastLine; }
            set
            {
                justifyLastLine = value;
                if (range != null) range.ProcessCells(c => c.CellStyle.Alignment.justifyLastLine = value);
            }
        }

        private XLAlignmentReadingOrderValues readingOrder;
        public XLAlignmentReadingOrderValues ReadingOrder
        {
            get { return readingOrder; }
            set
            {
                readingOrder = value;
                if (range != null) range.ProcessCells(c => c.CellStyle.Alignment.readingOrder = value);
            }
        }

        private Int32 relativeIndent;
        public Int32 RelativeIndent
        {
            get { return relativeIndent; }
            set
            {
                relativeIndent = value;
                if (range != null) range.ProcessCells(c => c.CellStyle.Alignment.relativeIndent = value);
            }
        }

        private Boolean shrinkToFit;
        public Boolean ShrinkToFit
        {
            get { return shrinkToFit; }
            set
            {
                shrinkToFit = value;
                if (range != null) range.ProcessCells(c => c.CellStyle.Alignment.shrinkToFit = value);
            }
        }

        private UInt32 textRotation;
        public UInt32 TextRotation
        {
            get
            {
                return textRotation;
            }
            set
            {
                if (value > 180)
                    throw new ArgumentException("TextRotation degree cannot be greater than 180.");
                else
                {
                    textRotation = value;
                    if (range != null) range.ProcessCells(c => c.CellStyle.Alignment.textRotation = value);
                }
            }
        }

        private Boolean wrapText;
        public Boolean WrapText
        {
            get { return wrapText; }
            set
            {
                wrapText = value;
                if (range != null) range.ProcessCells(c => c.CellStyle.Alignment.wrapText = value);
            }
        }

        public Boolean TopToBottom
        {
            get
            {
                return textRotation == 255;
            }
            set
            {
                if (value)
                {
                    textRotation = 255;
                    if (range != null) range.ProcessCells(c => c.CellStyle.Alignment.textRotation = 255);
                }
                else
                {
                    textRotation = 0;
                    if (range != null) range.ProcessCells(c => c.CellStyle.Alignment.textRotation = 0);
                }

            }
        }

        #endregion

        #region Constructors

        public XLAlignment(XLAlignment defaultAlignment, XLRange range)
        {
            this.range = range;
            if (defaultAlignment != null)
            {
                horizontal = defaultAlignment.horizontal;
                Vertical = defaultAlignment.Vertical;
                indent = defaultAlignment.indent;
                JustifyLastLine = defaultAlignment.JustifyLastLine;
                ReadingOrder = defaultAlignment.ReadingOrder;
                RelativeIndent = defaultAlignment.RelativeIndent;
                ShrinkToFit = defaultAlignment.ShrinkToFit;
                textRotation = defaultAlignment.textRotation;
                WrapText = defaultAlignment.WrapText;
            }
        }

        #endregion

        #region Overridden

        public override string ToString()
        {
            return
                Horizontal.ToString()
                + "-" + Vertical.ToString()
                + "-" + Indent.ToString()
                + "-" + JustifyLastLine.ToString()
                + "-" + ReadingOrder.ToString()
                + "-" + RelativeIndent.ToString()
                + "-" + ShrinkToFit.ToString()
                + "-" + textRotation.ToString()
                + "-" + WrapText.ToString()
                ;
        }

        #endregion
    }
}
