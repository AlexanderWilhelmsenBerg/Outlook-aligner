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
    private int _value;

    [Benchmark]
    public int Increment() => ++_value;
}
