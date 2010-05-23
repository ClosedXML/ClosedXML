using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel.Style
{
    public enum XLFillPatternValues
    {
        DarkDown,
        DarkGray,
        DarkGrid,
        DarkHorizontal,
        DarkTrellis,
        DarkUp,
        DarkVertical,
        Gray0625,
        Gray125,
        LightDown,
        LightGray,
        LightGrid,
        LightHorizontal,
        LightTrellis,
        LightUp,
        LightVertical,
        MediumGray,
        None,
        Solid
    }

    public class XLFill
    {
        #region Properties

        private XLRange range;

        public String BackgroundColor
        {
            get
            {
                return patternColor;
            }
            set
            {
                patternType = XLFillPatternValues.Solid;
                patternColor = value;
                patternBackgroundColor = value;
                if (range != null) range.ProcessCells(c =>
                {
                    c.CellStyle.Fill.patternType = XLFillPatternValues.Solid;
                    c.CellStyle.Fill.patternColor = value;
                    c.CellStyle.Fill.patternBackgroundColor = value;
                });
            }
        }


        private String patternColor;
        public String PatternColor
        {
            get
            {
                return patternColor;
            }
            set
            {
                this.patternColor = value;
                if (range != null) range.ProcessCells(c => c.CellStyle.Fill.patternColor = value);
            }
        }

        private String patternBackgroundColor;
        public String PatternBackgroundColor
        {
            get
            {
                return patternBackgroundColor;
            }
            set
            {
                this.patternBackgroundColor = value;
                if (range != null) range.ProcessCells(c => c.CellStyle.Fill.patternBackgroundColor = value);
            }
        }

        private XLFillPatternValues patternType;
        public XLFillPatternValues PatternType
        {
            get
            {
                return patternType;
            }
            set
            {
                this.patternType = value;
                if (range != null) range.ProcessCells(c => c.CellStyle.Fill.patternType = value);
            }
        }

        #endregion

        #region Constructors

        public XLFill(XLFill defaultFill, XLRange range)
        {
            this.range = range;
            if (defaultFill != null)
            {
                PatternType = defaultFill.PatternType;
                PatternColor = defaultFill.PatternColor;
                PatternBackgroundColor = defaultFill.PatternBackgroundColor;
            }
        }

        #endregion

        #region Overridden

        public override string ToString()
        {
            return BackgroundColor.ToString() + "-" + PatternType.ToString() + "-" + PatternColor.ToString();
        }

        #endregion
    }
}
