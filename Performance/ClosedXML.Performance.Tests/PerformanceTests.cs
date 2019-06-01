using BenchmarkDotNet.Reports;
using BenchmarkDotNet.Running;
using ClosedXML.Performance.Tests.Benchmarks;
using NUnit.Framework;
using System;
using System.Linq;

namespace ClosedXML.Performance.Tests
{
    public class PerformanceTests
    {
        /// <summary>
        /// For start, let's assume the performance should not degrade for more that 10%.
        /// </summary>
        private const double Tolerance = 0.10;

        [TestCase(typeof(CreateAndSaveWorkbook))]
        //Add other classes here
        public void CreateAndSaveWorkbook(Type benchmarkClass)
        {
            var res = BenchmarkRunner.Run(benchmarkClass);
            AssertBenchmarkResultsFitTolerance(res);
        }

        #region Private Methods

        private void AssertBenchmarkResultsFitTolerance(Summary benchmarkResult)
        {
            Assert.False(benchmarkResult.HasCriticalValidationErrors,
                string.Join(Environment.NewLine, benchmarkResult.ValidationErrors));

            Assert.True(benchmarkResult.Reports.All(r => r.Success), "Not all of the runs finished successfully");

            var baseLineCase = benchmarkResult.BenchmarksCases.Single(c => c.Descriptor.Baseline);
            var baseReport = benchmarkResult.Reports.Single(r => r.BenchmarkCase == baseLineCase);
            var otherReports = benchmarkResult.Reports.Where(r => r != baseReport);

            foreach (var report in otherReports)
            {
                var currentMeanValue = report.ResultStatistics.Mean;
                var baseMeanValue = baseReport.ResultStatistics.Mean;

                Assert.LessOrEqual(currentMeanValue, baseMeanValue * (1.0 + Tolerance),
                    $"The base execution time was {baseMeanValue} ns, the current is {currentMeanValue} ns.");

                foreach (var metricsKey in report.Metrics.Keys)
                {
                    var value = report.Metrics[metricsKey].Value;
                    var baseValue = baseReport.Metrics[metricsKey].Value;
                    var metricsName = report.Metrics[metricsKey].Descriptor.DisplayName;
                    if (report.Metrics[metricsKey].Descriptor.TheGreaterTheBetter)
                    {
                        Assert.GreaterOrEqual(value, baseValue * (1.0 - Tolerance),
                            $"The base value of {metricsName} was {baseValue}, the current is {value}. The greater the better.");
                    }
                    else
                    {
                        Assert.LessOrEqual(value, baseValue * (1.0 + Tolerance),
                            $"The base value of {metricsName} was {baseValue}, the current is {value}. The lower the better.");
                    }
                }
            }
        }

        #endregion Private Methods
    }
}
