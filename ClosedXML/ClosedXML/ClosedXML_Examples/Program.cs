using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML_Examples.Styles;

namespace ClosedXML_Examples
{
    class Program
    {
        static void Main(string[] args)
        {
            new HelloWorld().Create(@"c:\HelloWorld.xlsx");
            new BasicTable().Create(@"c:\BasicTable.xlsx");
            new StyleExamples().Create();
        }
    }
}