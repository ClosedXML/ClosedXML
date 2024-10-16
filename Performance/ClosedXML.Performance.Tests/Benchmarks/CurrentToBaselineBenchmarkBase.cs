using BenchmarkDotNet.Attributes;

namespace ClosedXML.Performance.Tests.Benchmarks
{
    public abstract class CurrentToBaselineBenchmarkBase
    {
        private readonly IPerformanceRunner _baseRunner;
        private readonly IPerformanceRunner _currentRunner;

        public CurrentToBaselineBenchmarkBase()
        {
            _baseRunner = new BaseLine.PerformanceRunner();
            _currentRunner = new Current.PerformanceRunner();
        }

        [Benchmark(Baseline = true)]
        public void Base()
        {
            TestMethod(_baseRunner);
        }

        [Benchmark]
        public void Current()
        {
            TestMethod(_currentRunner);
        }

        protected abstract void TestMethod(IPerformanceRunner performanceRunner);
    }
}
