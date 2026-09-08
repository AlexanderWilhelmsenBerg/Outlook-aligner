namespace OutlookAligner.OutlookHost.Forwarding;

internal sealed record ForwardCandidateInfo(
    string FolderName,
    string EntryId,
    string MessageClass,
    DateTime StartLocal,
    DateTime EndLocal,
    string? Subject);

internal sealed record ForwardCommandProbeInfo(
    string CommandId,
    bool IdentifierValid,
    string? Label,
    bool Visible,
    bool Enabled);

internal sealed record ForwardSpikeResult(
    int ExitCode,
    string Summary,
    string SourceSmtp,
    string GlobalAppointmentId,
    ForwardSpikeAction Action,
    string? Recipient,
    IReadOnlyList<string> SearchedFolders,
    IReadOnlyList<ForwardCandidateInfo> Candidates,
    ForwardCommandProbeInfo? CommandProbe,
    IReadOnlyList<string> Warnings)
{
    internal bool Success => ExitCode == 0;
}
