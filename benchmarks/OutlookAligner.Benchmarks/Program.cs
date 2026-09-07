using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Running;

namespace OutlookAligner.Benchmarks;

public static class Program
{
    public static void Main(string[] args)
    {
        BenchmarkRunner.Run<BootstrapBenchmark>(args: args);
    }
}

public class BootstrapBenchmark
{
    [Benchmark]
    public Guid CreateGuid() => Guid.NewGuid();
}
