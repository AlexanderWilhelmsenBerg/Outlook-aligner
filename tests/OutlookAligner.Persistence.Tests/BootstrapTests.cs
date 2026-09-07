using OutlookAligner.Persistence;

namespace OutlookAligner.Persistence.Tests;

public sealed class BootstrapTests
{
    [Fact]
    public void PersistenceAssemblyIsLoadable()
    {
        Assert.NotNull(typeof(PersistenceAssemblyMarker).Assembly);
    }
}
