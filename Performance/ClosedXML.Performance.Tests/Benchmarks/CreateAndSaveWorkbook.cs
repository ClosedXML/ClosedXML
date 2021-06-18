namespace ClosedXML.Performance.Tests.Benchmarks
{
    public class CreateAndSaveWorkbook : CurrentToBaselineBenchmarkBase
    {
        protected override void TestMethod(IPerformanceRunner performanceRunner) =>
            performanceRunner.CreateAndSaveEmptyWorkbook();
    }
}
