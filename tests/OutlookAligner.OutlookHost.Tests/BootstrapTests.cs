using OutlookAligner.Outlook.Contracts;

namespace OutlookAligner.OutlookHost.Tests;

public sealed class BootstrapTests
{
    [Fact]
    public void ProtocolVersionStartsAtOne()
    {
        Assert.Equal(1, OutlookHostProtocol.Version);
    }
}
