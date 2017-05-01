using System;

namespace ClosedXML_Sandbox
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            //Console.WriteLine("Running {0}", "OpenTestFile");
            //PerformanceRunner.TimeAction(PerformanceRunner.OpenTestFile);
            //Console.WriteLine();

            Console.WriteLine("Running {0}", "ReadComments");
            PerformanceRunner.ReadComments();
            Console.WriteLine();

            //Console.WriteLine("Running {0}", "RunInsertTable");
            //PerformanceRunner.TimeAction(PerformanceRunner.RunInsertTable);
            //Console.WriteLine();

            Console.WriteLine("Press any key to continue");
            Console.ReadKey();
        }
    }
}