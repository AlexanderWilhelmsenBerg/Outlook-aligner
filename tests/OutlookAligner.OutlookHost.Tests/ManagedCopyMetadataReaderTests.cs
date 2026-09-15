using OutlookAligner.Outlook.Contracts;
using OutlookAligner.OutlookHost.Calendars;
using Xunit;

namespace OutlookAligner.OutlookHost.Tests;

public sealed class ManagedCopyMetadataReaderTests
{
    [Fact]
    public void EmptyMetadataIsNotManaged()
    {
        var metadata = ManagedCopyMetadataReader.Classify(null, null, null, null, null);

        Assert.Equal(ManagedCopyState.None, metadata.State);
    }

    [Fact]
    public void CompleteCurrentSchemaMetadataIsValid()
    {
        var metadata = ManagedCopyMetadataReader.Classify(
            Guid.NewGuid().ToString("D"),
            "global-id",
            "source@example.invalid",
            OutlookAlignerManagedProperties.CopyTypeFull,
            OutlookAlignerManagedProperties.CurrentSchemaVersion);

        Assert.Equal(ManagedCopyState.Valid, metadata.State);
    }

    [Fact]
    public void PartialMetadataIsNeverTreatedAsManaged()
    {
        var metadata = ManagedCopyMetadataReader.Classify(
            Guid.NewGuid().ToString("D"),
            "global-id",
            sourceAccountId: null,
            OutlookAlignerManagedProperties.CopyTypeBusy,
            OutlookAlignerManagedProperties.CurrentSchemaVersion);

        Assert.Equal(ManagedCopyState.Incomplete, metadata.State);
    }

    [Fact]
    public void InvalidSyncGroupIdIsIncomplete()
    {
        var metadata = ManagedCopyMetadataReader.Classify(
            "not-a-guid",
            "global-id",
            "source@example.invalid",
            OutlookAlignerManagedProperties.CopyTypeFull,
            OutlookAlignerManagedProperties.CurrentSchemaVersion);

        Assert.Equal(ManagedCopyState.Incomplete, metadata.State);
    }

    [Fact]
    public void UnknownCopyTypeIsIncomplete()
    {
        var metadata = ManagedCopyMetadataReader.Classify(
            Guid.NewGuid().ToString("D"),
            "global-id",
            "source@example.invalid",
            "SomethingElse",
            OutlookAlignerManagedProperties.CurrentSchemaVersion);

        Assert.Equal(ManagedCopyState.Incomplete, metadata.State);
    }

    [Fact]
    public void FutureSchemaIsUnsupportedRatherThanManaged()
    {
        var metadata = ManagedCopyMetadataReader.Classify(
            Guid.NewGuid().ToString("D"),
            "global-id",
            "source@example.invalid",
            OutlookAlignerManagedProperties.CopyTypeFull,
            "999");

        Assert.Equal(ManagedCopyState.UnsupportedSchema, metadata.State);
    }
}
