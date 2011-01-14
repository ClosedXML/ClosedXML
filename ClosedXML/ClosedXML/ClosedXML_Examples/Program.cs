using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML_Examples.Styles;
using ClosedXML_Examples.Columns;
using ClosedXML_Examples.Rows;
using ClosedXML_Examples.Misc;
using ClosedXML_Examples.Ranges;
using ClosedXML_Examples.PageSetup;

namespace ClosedXML_Examples
{
    public class Program
    {
        static void Main(string[] args)
        {
            CreateFiles.CreateAllFiles();
            LoadFiles.LoadAllFiles();
        }

        public static void ExecuteMain()
        {
            CreateFiles.CreateAllFiles();
            LoadFiles.LoadAllFiles();
        }
    }
}