using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace ClosedXML.Excel
{
    public interface xIXLWorkbook
    {
        IXLWorksheets Worksheets { get; }
        String Name { get; }
        String FullName { get; }
        void SaveAs(String file, Boolean overwrite = false);
        void Load(String file);
    }
}
