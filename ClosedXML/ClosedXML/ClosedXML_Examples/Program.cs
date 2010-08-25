using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML_Examples.Styles;
using ClosedXML_Examples.Columns;
using ClosedXML_Examples.Rows;

namespace ClosedXML_Examples
{
    class Program
    {
        static void Main(string[] args)
        {
            new HelloWorld().Create(@"c:\HelloWorld.xlsx");
            new BasicTable().Create(@"c:\BasicTable.xlsx");
            new StyleExamples().Create();
            new ColumnSettings().Create(@"c:\ColumnSettings.xlsx");
            new RowSettings().Create(@"c:\RowSettings.xlsx");
        }
    }
}