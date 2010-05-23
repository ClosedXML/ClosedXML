using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel.Style
{
    public enum XLBorderStyleValues
    {
        DashDot,
        DashDotDot,
        Dashed,
        Dotted,
        Double,
        Hair,
        Medium,
        MediumDashDot,
        MediumDashDotDot,
        MediumDashed,
        None,
        SlantDashDot,
        Thick,
        Thin
    }

    public class XLBorder
    {
        #region Properties

        private XLRange range;

        private XLBorderStyleValues leftBorder;
        public XLBorderStyleValues LeftBorder
        {
            get { return leftBorder; }
            set
            {
                leftBorder = value;
                if (range != null) range.ProcessCells(c => c.CellStyle.Border.leftBorder = value);
            }
        }

        private String leftBorderColor;
        public String LeftBorderColor
        {
            get { return leftBorderColor; }
            set
            {
                leftBorderColor = value;
                if (range != null) range.ProcessCells(c => c.CellStyle.Border.leftBorderColor = value);
            }
        }

        private XLBorderStyleValues rightBorder;
        public XLBorderStyleValues RightBorder
        {
            get { return rightBorder; }
            set
            {
                rightBorder = value;
                if (range != null) range.ProcessCells(c => c.CellStyle.Border.rightBorder = value);
            }
        }

        private String rightBorderColor;
        public String RightBorderColor
        {
            get { return rightBorderColor; }
            set
            {
                rightBorderColor = value;
                if (range != null) range.ProcessCells(c => c.CellStyle.Border.rightBorderColor = value);
            }
        }

        private XLBorderStyleValues topBorder;
        public XLBorderStyleValues TopBorder
        {
            get { return topBorder; }
            set
            {
                topBorder = value;
                if (range != null) range.ProcessCells(c => c.CellStyle.Border.topBorder = value);
            }
        }

        private String topBorderColor;
        public String TopBorderColor
        {
            get { return topBorderColor; }
            set
            {
                topBorderColor = value;
                if (range != null) range.ProcessCells(c => c.CellStyle.Border.topBorderColor = value);
            }
        }

        private XLBorderStyleValues bottomBorder;
        public XLBorderStyleValues BottomBorder
        {
            get { return bottomBorder; }
            set
            {
                bottomBorder = value;
                if (range != null) range.ProcessCells(c => c.CellStyle.Border.bottomBorder = value);
            }
        }

        private String bottomBorderColor;
        public String BottomBorderColor
        {
            get { return bottomBorderColor; }
            set
            {
                bottomBorderColor = value;
                if (range != null) range.ProcessCells(c => c.CellStyle.Border.bottomBorderColor = value);
            }
        }

        private Boolean diagonalUp;
        public Boolean DiagonalUp
        {
            get { return diagonalUp; }
            set
            {
                diagonalUp = value;
                if (range != null) range.ProcessCells(c => c.CellStyle.Border.diagonalUp = value);
            }
        }

        private Boolean diagonalDown;
        public Boolean DiagonalDown
        {
            get { return diagonalDown; }
            set
            {
                diagonalDown = value;
                if (range != null) range.ProcessCells(c => c.CellStyle.Border.diagonalDown = value);
            }
        }

        private XLBorderStyleValues diagonalBorder;
        public XLBorderStyleValues DiagonalBorder
        {
            get { return diagonalBorder; }
            set
            {
                diagonalBorder = value;
                if (range != null) range.ProcessCells(c => c.CellStyle.Border.diagonalBorder = value);
            }
        }

        private String diagonalBorderColor;
        public String DiagonalBorderColor
        {
            get { return diagonalBorderColor; }
            set
            {
                diagonalBorderColor = value;
                if (range != null) range.ProcessCells(c => c.CellStyle.Border.diagonalBorderColor = value);
            }
        }

        #endregion

        #region Constructors

        public XLBorder(XLBorder defaultBorder, XLRange range)
        {
            this.range = range;
            if (defaultBorder != null)
            {
                LeftBorder = defaultBorder.LeftBorder;
                LeftBorderColor = defaultBorder.LeftBorderColor;
                RightBorder = defaultBorder.RightBorder;
                RightBorderColor = defaultBorder.RightBorderColor;
                TopBorder = defaultBorder.TopBorder;
                TopBorderColor = defaultBorder.TopBorderColor;
                BottomBorder = defaultBorder.BottomBorder;
                BottomBorderColor = defaultBorder.BottomBorderColor;
                DiagonalBorder = defaultBorder.DiagonalBorder;
                DiagonalBorderColor = defaultBorder.DiagonalBorderColor;
                DiagonalUp = defaultBorder.DiagonalUp;
                DiagonalDown = defaultBorder.DiagonalDown;
            }
        }

        #endregion

        #region Overridden

        public override string ToString()
        {
            return
                LeftBorder.ToString() + "-" +
                LeftBorderColor.ToString() + "-" +
                RightBorder.ToString() + "-" +
                RightBorderColor.ToString() + "-" +
                TopBorder.ToString() + "-" +
                TopBorderColor.ToString() + "-" +
                BottomBorder.ToString() + "-" +
                BottomBorderColor.ToString() + "-" +
                DiagonalBorder.ToString() + "-" +
                DiagonalBorderColor.ToString() + "-" +
                DiagonalUp.ToString() + "-" +
                DiagonalDown.ToString();

        }

        #endregion
    }
}
