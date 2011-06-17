using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML_Examples.Delete;
using ClosedXML_Examples.Styles;
using ClosedXML_Examples.Columns;
using ClosedXML_Examples.Rows;
using ClosedXML_Examples.Misc;
using ClosedXML_Examples.Ranges;
using ClosedXML_Examples.PageSetup;

namespace ClosedXML_Examples
{
    public class ModifyFiles
    {
        public static void Run()
        {
            new DeleteRows().Create(@"C:\Excel Files\Modify\DeleteRows.xlsx");
        }
    }
}
