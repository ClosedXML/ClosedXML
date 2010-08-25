using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel.Style;

namespace ClosedXML.Excel
{
    public interface IXLWorkbook
    {
        IXLWorksheets Worksheets { get; }
        String Name { get; }
        String FullName { get; }
        void SaveAs(String file, Boolean overwrite = false);

    }
}
