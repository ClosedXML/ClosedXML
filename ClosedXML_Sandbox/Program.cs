using System;

namespace ClosedXML_Sandbox
{
    class Program
    {
        private static void Main(string[] args)
        {
            PerformanceRunner.TimeAction(PerformanceRunner.RunInsertTable);

            Console.WriteLine("Press any key to continue");
            Console.ReadKey();
        }
    }
}