namespace ClosedXML.Tests.Performance;

[TestFixture]
public class BenchmarkTestRunner
{
    [Test, Explicit("Performance test")]
    public void WorkbookOperationPerformanceTests()
    {
        BenchmarkRunnerHelper.RunBenchmark<WorkbookOperationBenchmarks>();
    }
}