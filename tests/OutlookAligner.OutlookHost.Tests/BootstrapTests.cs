using OutlookAligner.Outlook.Contracts;
using Xunit;

namespace OutlookAligner.OutlookHost.Tests;

public sealed class BootstrapTests
{
    [Fact]
    public void ProtocolVersionStartsAtOne()
    {
        Assert.Equal(1, OutlookHostProtocol.Version);
    }
}
