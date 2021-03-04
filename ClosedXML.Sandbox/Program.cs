using System;

namespace ClosedXML.Sandbox
{
    internal static class Program
    {
        private static void Main(string[] args)
        {
            Console.WriteLine("Running {0}", nameof(PerformanceRunner.OpenTestFile));
            PerformanceRunner.TimeAction(PerformanceRunner.OpenTestFile);
            Console.WriteLine();

            // Disable this block by default - I don't use it often
#if false

            Console.WriteLine("Running {0}", nameof(PerformanceRunner.RunInsertTable));
            PerformanceRunner.TimeAction(PerformanceRunner.RunInsertTable);
            Console.WriteLine();

            Console.WriteLine("Running {0}", nameof(PerformanceRunner.PerformHeavyCalculation));
            PerformanceRunner.TimeAction(PerformanceRunner.PerformHeavyCalculation);
            Console.WriteLine();
#endif

            Console.WriteLine("Press any key to continue");
            Console.ReadKey();
        }
    }
}
