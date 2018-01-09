using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace AppForIssueTracking
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var wb = new XLWorkbook("TestSumIf.xlsx"))
            {
                // Issue 1 - ExpressionParseException thrown
                var ws = wb.Worksheet("SUMIF Not Same Columns");
                var value = ws.Cell("E1").Value;
                Console.WriteLine(value);
                // Issue 2 - Evaluation hangs
                //var ws = wb.Worksheet("SUMIF Same Columns");
                //var value = ws.Cell("B7").Value;
                //Console.WriteLine(value);
            }

        }
    }
}
