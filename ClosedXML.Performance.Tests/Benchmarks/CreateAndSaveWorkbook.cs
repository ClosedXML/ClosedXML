using BenchmarkDotNet.Attributes;


namespace ClosedXML.Performance.Tests.Benchmarks
{

    public class CreateAndSaveWorkbook
    {
        private readonly IPerformanceRunner _baseRunner;
        private readonly IPerformanceRunner _currentRunner;

        public CreateAndSaveWorkbook()
        {
            _baseRunner = new BaseLine.PerformanceRunner();
            _currentRunner = new Current.PerformanceRunner();
        }

        [Benchmark(Baseline = true)]
        public void Base()
        {
            _baseRunner.CreateAndSaveEmptyWorkbook();
        }


        [Benchmark]
        public void Current()
        {
            _currentRunner.CreateAndSaveEmptyWorkbook();
        }
    }
}
