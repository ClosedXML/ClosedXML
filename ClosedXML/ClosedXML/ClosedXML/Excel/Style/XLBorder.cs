using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
namespace ClosedXML.Excel
{
    internal class XLBorder : IXLBorder
    {
        IXLStylized container;

        public XLBorder() : this(null, XLWorkbook.DefaultStyle.Border) { }

        public XLBorder(IXLStylized container, IXLBorder defaultBorder)
        {
            this.container = container;
            if (defaultBorder != null)
            {
                leftBorder = defaultBorder.LeftBorder;
                leftBorderColor = defaultBorder.LeftBorderColor;
                rightBorder = defaultBorder.RightBorder;
                rightBorderColor = defaultBorder.RightBorderColor;
                topBorder = defaultBorder.TopBorder;
                topBorderColor = defaultBorder.TopBorderColor;
                bottomBorder = defaultBorder.BottomBorder;
                bottomBorderColor = defaultBorder.BottomBorderColor;
                diagonalBorder = defaultBorder.DiagonalBorder;
                diagonalBorderColor = defaultBorder.DiagonalBorderColor;
                diagonalUp = defaultBorder.DiagonalUp;
                diagonalDown = defaultBorder.DiagonalDown;
            }
        }

        public XLBorderStyleValues OutsideBorder
        {
            set
            {
                if (container != null && !container.UpdatingStyle)
                {
                    foreach (var r in container.RangesUsed)
                    {
                        r.FirstColumn().Style.Border.LeftBorder = value;
                        r.LastColumn().Style.Border.RightBorder = value;
                        r.FirstRow().Style.Border.TopBorder = value;
                        r.LastRow().Style.Border.BottomBorder = value;
                    }
                }
            }
        }

        public IXLColor OutsideBorderColor
        {
            set
            {
                if (container != null && !container.UpdatingStyle)
                {
                    foreach (var r in container.RangesUsed)
                    {
                        r.FirstColumn().Style.Border.LeftBorderColor = value;
                        r.LastColumn().Style.Border.RightBorderColor = value;
                        r.FirstRow().Style.Border.TopBorderColor = value;
                        r.LastRow().Style.Border.BottomBorderColor = value;
                    }
                }
            }
        }

        private XLBorderStyleValues leftBorder;
        public XLBorderStyleValues LeftBorder
        {
            get
            {
                return leftBorder;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Border.LeftBorder = value);
                else
                    leftBorder = value;
            }
        }

        private IXLColor leftBorderColor;
        public IXLColor LeftBorderColor
        {
            get
            {
                return leftBorderColor;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Border.LeftBorderColor = value);
                else
                    leftBorderColor = value;
            }
        }

        private XLBorderStyleValues rightBorder;
        public XLBorderStyleValues RightBorder
        {
            get
            {
                return rightBorder;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Border.RightBorder = value);
                else
                    rightBorder = value;
            }
        }

        private IXLColor rightBorderColor;
        public IXLColor RightBorderColor
        {
            get
            {
                return rightBorderColor;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Border.RightBorderColor = value);
                else
                    rightBorderColor = value;
            }
        }

        private XLBorderStyleValues topBorder;
        public XLBorderStyleValues TopBorder
        {
            get
            {
                return topBorder;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Border.TopBorder = value);
                else
                    topBorder = value;
            }
        }

        private IXLColor topBorderColor;
        public IXLColor TopBorderColor
        {
            get
            {
                return topBorderColor;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Border.TopBorderColor = value);
                else
                    topBorderColor = value;
            }
        }

        private XLBorderStyleValues bottomBorder;
        public XLBorderStyleValues BottomBorder
        {
            get
            {
                return bottomBorder;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Border.BottomBorder = value);
                else
                    bottomBorder = value;
            }
        }

        private IXLColor bottomBorderColor;
        public IXLColor BottomBorderColor
        {
            get
            {
                return bottomBorderColor;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Border.BottomBorderColor = value);
                else
                    bottomBorderColor = value;
            }
        }

        private XLBorderStyleValues diagonalBorder;
        public XLBorderStyleValues DiagonalBorder
        {
            get
            {
                return diagonalBorder;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Border.DiagonalBorder = value);
                else
                    diagonalBorder = value;
            }
        }

        private IXLColor diagonalBorderColor;
        public IXLColor DiagonalBorderColor
        {
            get
            {
                return diagonalBorderColor;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Border.DiagonalBorderColor = value);
                else
                    diagonalBorderColor = value;
            }
        }

        private Boolean diagonalUp;
        public Boolean DiagonalUp
        {
            get
            {
                return diagonalUp;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Border.DiagonalUp = value);
                else
                    diagonalUp = value;
            }
        }

        private Boolean diagonalDown;
        public Boolean DiagonalDown
        {
            get
            {
                return diagonalDown;
            }
            set
            {
                if (container != null && !container.UpdatingStyle)
                    container.Styles.ForEach(s => s.Border.DiagonalDown = value);
                else
                    diagonalDown = value;
            }
        }

        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.Append(LeftBorder.ToString());
            sb.Append("-");
            sb.Append(LeftBorderColor.ToString());
            sb.Append("-");
            sb.Append(RightBorder.ToString());
            sb.Append("-");
            sb.Append(RightBorderColor.ToString());
            sb.Append("-");
            sb.Append(TopBorder.ToString());
            sb.Append("-");
            sb.Append(TopBorderColor.ToString());
            sb.Append("-");
            sb.Append(BottomBorder.ToString());
            sb.Append("-");
            sb.Append(BottomBorderColor.ToString());
            sb.Append("-");
            sb.Append(DiagonalBorder.ToString());
            sb.Append("-");
            sb.Append(DiagonalBorderColor.ToString());
            sb.Append("-");
            sb.Append(DiagonalUp.ToString());
            sb.Append("-");
            sb.Append(DiagonalDown.ToString());
            return sb.ToString();
        }

        public bool Equals(IXLBorder other)
        {
            return
                   this.LeftBorder.Equals(other.LeftBorder)
                && this.LeftBorderColor.Equals(other.LeftBorderColor)
                && this.RightBorder.Equals(other.RightBorder)
                && this.RightBorderColor.Equals(other.RightBorderColor)
                && this.TopBorder.Equals(other.TopBorder)
                && this.TopBorderColor.Equals(other.TopBorderColor)
                && this.BottomBorder.Equals(other.BottomBorder)
                && this.BottomBorderColor.Equals(other.BottomBorderColor)
                && this.DiagonalBorder.Equals(other.DiagonalBorder)
                && this.DiagonalBorderColor.Equals(other.DiagonalBorderColor)
                && this.DiagonalUp.Equals(other.DiagonalUp)
                && this.DiagonalDown.Equals(other.DiagonalDown)
                ;
        }

        public override bool Equals(object obj)
        {
            return this.Equals((XLBorder)obj);
        }

        public override int GetHashCode()
        {
                return (Int32)LeftBorder
                    ^ LeftBorderColor.GetHashCode()
                    ^ (Int32)RightBorder
                    ^ RightBorderColor.GetHashCode()
                    ^ (Int32)TopBorder
                    ^ TopBorderColor.GetHashCode()
                    ^ (Int32)BottomBorder
                    ^ BottomBorderColor.GetHashCode()
                    ^ (Int32)DiagonalBorder
                    ^ DiagonalBorderColor.GetHashCode()
                    ^ DiagonalUp.GetHashCode()
                    ^ DiagonalDown.GetHashCode();
        }

        public IXLStyle SetOutsideBorder(XLBorderStyleValues value) { OutsideBorder = value; return container.Style; }
        public IXLStyle SetOutsideBorderColor(IXLColor value) { OutsideBorderColor = value; return container.Style; }

        public IXLStyle SetLeftBorder(XLBorderStyleValues value) { LeftBorder = value; return container.Style; }
        public IXLStyle SetLeftBorderColor(IXLColor value) { LeftBorderColor = value; return container.Style; }
        public IXLStyle SetRightBorder(XLBorderStyleValues value) { RightBorder = value; return container.Style; }
        public IXLStyle SetRightBorderColor(IXLColor value) { RightBorderColor = value; return container.Style; }
        public IXLStyle SetTopBorder(XLBorderStyleValues value) { TopBorder = value; return container.Style; }
        public IXLStyle SetTopBorderColor(IXLColor value) { TopBorderColor = value; return container.Style; }
        public IXLStyle SetBottomBorder(XLBorderStyleValues value) { BottomBorder = value; return container.Style; }
        public IXLStyle SetBottomBorderColor(IXLColor value) { BottomBorderColor = value; return container.Style; }
        public IXLStyle SetDiagonalUp() { DiagonalUp = true; return container.Style; }	public IXLStyle SetDiagonalUp(Boolean value) { DiagonalUp = value; return container.Style; }
        public IXLStyle SetDiagonalDown() { DiagonalDown = true; return container.Style; }	public IXLStyle SetDiagonalDown(Boolean value) { DiagonalDown = value; return container.Style; }
        public IXLStyle SetDiagonalBorder(XLBorderStyleValues value) { DiagonalBorder = value; return container.Style; }
        public IXLStyle SetDiagonalBorderColor(IXLColor value) { DiagonalBorderColor = value; return container.Style; }

    }
}
