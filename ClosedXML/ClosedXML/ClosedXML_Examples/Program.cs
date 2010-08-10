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
            //var helloWorld = new HelloWorld();
            //helloWorld.Create(@"c:\HelloWorld.xlsx");

            //new StyleFont().Create(@"c:\styleFont.xlsx");

            //new StyleFill().Create(@"c:\styleFill.xlsx");

            new StyleBorder().Create(@"c:\styleBorder.xlsx");

            //var basicTable = new BasicTable();
            //basicTable.Create(@"c:\BasicTable.xlsx");
        }
    }
}
