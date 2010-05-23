using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML_Examples
{
    class Program
    {
        static void Main(string[] args)
        {
            var helloWorld = new HelloWorld();
            helloWorld.Create(@"c:\HelloWorld.xlsx");

            var basicTable = new BasicTable();
            basicTable.Create(@"c:\BasicTable.xlsx");
        }
    }
}
