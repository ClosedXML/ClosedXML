using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLChart: IXLChart
    {
        public XLChart()
        {
            Position = new XLDrawingPosition();
        }
        IXLDrawingPosition Position { get; set; }

        IXLDrawingPosition IXLChart.Position
        {
            get { throw new NotImplementedException(); }
        }
    }
}
