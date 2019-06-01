using System;
using System.Linq;
using BenchmarkDotNet.Reports;
using BenchmarkDotNet.Running;
using ClosedXML.Performance.Tests.Benchmarks;
using NUnit.Framework;

namespace ClosedXML.Performance.Tests
{
    public class PerformanceTests
    {
        private double Tolerance = 0.05;

        [Test]
        public void CreateAndSaveWorkbook()
        {
            var res = BenchmarkRunner.Run<CreateAndSaveWorkbook>();

            Assert.False(res.HasCriticalValidationErrors, string.Join(Environment.NewLine, res.ValidationErrors));
            Assert.True(res.Reports.All(r => r.Success), "Not all of the runs finished successfully");

            var baseLineCase = res.BenchmarksCases.Single(c => c.Descriptor.Baseline);
            var baseReport = res.Reports.Single(r => r.BenchmarkCase == baseLineCase);
            var otherReports = res.Reports.Where(r => r != baseReport);

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
    }
}
