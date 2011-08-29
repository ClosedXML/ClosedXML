using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLDrawingSize
    {
        Double Height { get; set; }
        Double Width { get; set; }
        Double ScaleHeight { get; set; }
        Double ScaleWidth { get; set; }
        Boolean LockAspectRatio { get; set; }

        IXLDrawingStyle SetHeight(Double value);
        IXLDrawingStyle SetWidth(Double value);
        IXLDrawingStyle SetScaleHeight(Double value);
        IXLDrawingStyle SetScaleWidth(Double value);
        IXLDrawingStyle SetLockAspectRatio(); IXLDrawingStyle SetLockAspectRatio(Boolean value);

    }
}
