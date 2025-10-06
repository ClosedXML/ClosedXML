using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Loggers;
using BenchmarkDotNet.Running;

namespace ClosedXML.Tests.Performance;

public static class BenchmarkRunnerHelper
{
    public static void RunBenchmark<T>() where T : class
    {
        var logger = new AccumulationLogger();

        var config = ManualConfig.Create(DefaultConfig.Instance)
            .AddLogger(logger)
            .WithOptions(ConfigOptions.DisableOptimizationsValidator);

        BenchmarkRunner.Run<T>(config);

        TestContext.Out.WriteLine(logger.GetLog());
    }
}