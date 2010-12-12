using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLAlignment : IXLAlignment
    {
        IXLStylized container;

        public XLAlignment() : this(null, XLWorkbook.DefaultStyle.Alignment) { }

        public XLAlignment(IXLStylized container, IXLAlignment d = null)
        {
            this.container = container;
            if (d != null)
            {
                horizontal = d.Horizontal;
                vertical = d.Vertical;
                indent = d.Indent;
                justifyLastLine = d.JustifyLastLine;
                readingOrder = d.ReadingOrder;
                relativeIndent = d.RelativeIndent;
                shrinkToFit = d.ShrinkToFit;
                textRotation = d.TextRotation;
                wrapText = d.WrapText;
            }
        }

        private XLAlignmentHorizontalValues horizontal;
        public XLAlignmentHorizontalValues Horizontal
        {
            get
            {
                return horizontal;
            }
            set
            {
                Boolean updateIndent = !(
                    value == XLAlignmentHorizontalValues.Left
                    || value == XLAlignmentHorizontalValues.Right
                    || value == XLAlignmentHorizontalValues.Distributed
                    );

                if (container != null && !container.UpdatingStyle)
                {
                    container.Styles.ForEach(s => {
                        s.Alignment.Horizontal = value;
                        if (updateIndent)
                            s.Alignment.Indent = 0;
                    });
                }
                else
                {
                    horizontal = value;
                    if (updateIndent)
                        indent = 0;
                }
            }
        }

        private XLAlignmentVerticalValues vertical;
        public XLAlignmentVerticalValues Vertical
        {
            get
            {
                return vertical;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Alignment.Vertical = value);
                else
                    vertical = value;
            }
        }

        private Int32 indent;
        public Int32 Indent
        {
            get
            {
                return indent;
            }
            set
            {
                if (Horizontal == XLAlignmentHorizontalValues.General)
                    Horizontal = XLAlignmentHorizontalValues.Left;

                if (value > 0 && !(
                    Horizontal == XLAlignmentHorizontalValues.Left
                    || Horizontal == XLAlignmentHorizontalValues.Right
                    || Horizontal == XLAlignmentHorizontalValues.Distributed
                    ))
                {
                    throw new ArgumentException("For indents, only left, right, and distributed horizontal alignments are supported.");
                }

                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Alignment.Indent = value);
                else
                    indent = value;
            }
        }

        private Boolean justifyLastLine;
        public Boolean JustifyLastLine
        {
            get
            {
                return justifyLastLine;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Alignment.JustifyLastLine = value);
                else
                    justifyLastLine = value;
            }
        }

        private XLAlignmentReadingOrderValues readingOrder;
        public XLAlignmentReadingOrderValues ReadingOrder
        {
            get
            {
                return readingOrder;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Alignment.ReadingOrder = value);
                else
                    readingOrder = value;
            }
        }

        private Int32 relativeIndent;
        public Int32 RelativeIndent
        {
            get
            {
                return relativeIndent;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Alignment.RelativeIndent = value);
                else
                    relativeIndent = value;
            }
        }

        private Boolean shrinkToFit;
        public Boolean ShrinkToFit
        {
            get
            {
                return shrinkToFit;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Alignment.ShrinkToFit = value);
                else
                    shrinkToFit = value;
            }
        }

        private Int32 textRotation;
        public Int32 TextRotation
        {
            get
            {
                return textRotation;
            }
            set
            {
                if ( value != 255 && (value < 0 || value > 180) )
                    throw new ArgumentException("TextRotation must be between 0 and 180 degrees, or 255.");

                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Alignment.TextRotation = value);
                else
                    textRotation = value;
            }
        }

        private Boolean wrapText;
        public Boolean WrapText
        {
            get
            {
                return wrapText;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Alignment.WrapText = value);
                else
                    wrapText = value;
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
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Alignment.TextRotation = value ? 255 : 0 );
                else
                    textRotation = value ? 255 : 0;
            }
        }


        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.Append(Horizontal);
            sb.Append("-");
            sb.Append(Vertical);
            sb.Append("-");
            sb.Append(Indent);
            sb.Append("-");
            sb.Append(JustifyLastLine);
            sb.Append("-");
            sb.Append(ReadingOrder);
            sb.Append("-");
            sb.Append(RelativeIndent);
            sb.Append("-");
            sb.Append(ShrinkToFit);
            sb.Append("-");
            sb.Append(TextRotation);
            sb.Append("-");
            sb.Append(WrapText);
            sb.Append("-");
            return sb.ToString();
        }

        public override bool Equals(object obj)
        {
            return this.GetHashCode().Equals(obj.GetHashCode());
        }

        public override int GetHashCode()
        {
            unchecked // Overflow is fine, just wrap
            {
                int hash = 17;
                hash = hash * 23 + Horizontal.GetHashCode();
                hash = hash * 23 + Vertical.GetHashCode();
                hash = hash * 23 + Indent.GetHashCode();
                hash = hash * 23 + JustifyLastLine.GetHashCode();
                hash = hash * 23 + ReadingOrder.GetHashCode();
                hash = hash * 23 + RelativeIndent.GetHashCode();
                hash = hash * 23 + ShrinkToFit.GetHashCode();
                hash = hash * 23 + TextRotation.GetHashCode();
                hash = hash * 23 + WrapText.GetHashCode();
                return hash;
            }
        }

        public bool Equals(IXLAlignment other)
        {
            return 
               this.Horizontal.Equals(other.Horizontal)
            && this.Vertical.Equals(other.Vertical)
            && this.Indent.Equals(other.Indent)
            && this.JustifyLastLine.Equals(other.JustifyLastLine)
            && this.ReadingOrder.Equals(other.ReadingOrder)
            && this.RelativeIndent.Equals(other.RelativeIndent)
            && this.ShrinkToFit.Equals(other.ShrinkToFit)
            && this.TextRotation.Equals(other.TextRotation)
            && this.WrapText.Equals(other.WrapText)
            ;
        }

    }
}
