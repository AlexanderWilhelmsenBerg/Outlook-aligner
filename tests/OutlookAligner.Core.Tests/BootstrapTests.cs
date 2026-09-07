using OutlookAligner.Core;
using Xunit;

namespace OutlookAligner.Core.Tests;

public sealed class BootstrapTests
{
    [Fact]
    public void CoreAssemblyIsLoadable()
    {
        Assert.NotNull(typeof(CoreAssemblyMarker).Assembly);
    }
}
