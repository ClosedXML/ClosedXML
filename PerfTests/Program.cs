using System;
using BaseLine = ClosedXML.Performance.BaseLine;
using Current = ClosedXML.Performance.Current;

namespace PerfTests
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Running {0}", nameof(BaseLine.PerformanceRunner.Run1220));
            Console.Write("Base    ");
            BaseLine.PerformanceRunner.TimeAction(BaseLine.PerformanceRunner.Run1220);
            Console.WriteLine();
            Console.Write("Current ");
            Current.PerformanceRunner.TimeAction(Current.PerformanceRunner.Run1220);
            Console.WriteLine("------------");

            Console.WriteLine("Running {0}", nameof(BaseLine.PerformanceRunner.Run1221));
            Console.Write("Base    ");
            BaseLine.PerformanceRunner.TimeAction(BaseLine.PerformanceRunner.Run1221);
            Console.WriteLine();
            Console.Write("Current ");
            Current.PerformanceRunner.TimeAction(Current.PerformanceRunner.Run1221);
            Console.WriteLine("------------");            // Disable this block by default - I don't use it often
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
