using System;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace ClosedXML.Excel.Charts
{
    public class ChartLabel
    {
        public bool IsEnabled { get; set; }
        public int DecimalNumbers { get; set; }
        public string UnitSuffix { get; set; }
        public DataLabelPositionValues Position { get; set; }
        public LabelType LabelType { get; set; }
        public String HexColor { get; set; }
    }
}
