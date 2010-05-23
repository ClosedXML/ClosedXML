using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

namespace ClosedXML_Sandbox
{
    class Program
    {
        static void Main(string[] args)
        {
            var wbExample = new XLWorkbook(@"c:\Example.xlsx");
            var wsWorld = wbExample.Worksheets.Add("World");
            var wsNameList = wbExample.Worksheets.Add("Name List");
            var wsDeleteMe = wbExample.Worksheets.Add("Delete Me");

            

            var a1 = wsNameList.Cell("A1");
            var a2 = wsNameList.Cell("A2");
            a1.Value = "Hello!";
            wbExample.Save();
            // a2.Font.Bold = true;
        }
    }
}
