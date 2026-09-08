using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using OutlookAligner.Outlook.Contracts;

namespace OutlookAligner.App.Services;

internal sealed class OutlookHostProcessClient
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
    };

    internal async Task<OutlookProbeResult> ScanAsync(int days, CancellationToken cancellationToken = default)
    {
        var result = await RunAsync(
            ["--days", days.ToString(CultureInfo.InvariantCulture), "--json", "--include-details"],
            cancellationToken).ConfigureAwait(false);

        result.ThrowIfFailed("Outlook calendar scan");

        var probe = JsonSerializer.Deserialize<OutlookProbeResult>(result.StandardOutput, JsonOptions)
            ?? throw new InvalidOperationException("OutlookHost returned an empty calendar result.");

        if (probe.ProtocolVersion != OutlookHostProtocol.Version)
        {
            throw new InvalidOperationException(
                $"OutlookHost protocol {probe.ProtocolVersion} is incompatible with UI protocol {OutlookHostProtocol.Version}.");
        }

        return probe;
    }

    internal Task<HostCommandResult> ProbeForwardAsync(
        string sourceSmtp,
        string globalAppointmentId,
        string entryId,
        CancellationToken cancellationToken = default)
        => RunAsync(
            [
                "--forward-spike",
                "--source-smtp", sourceSmtp,
                "--global-id", globalAppointmentId,
                "--probe-calendar-command",
                "--entry-id", entryId,
                "--include-details",
            ],
            cancellationToken);

    internal Task<HostCommandResult> PrepareForwardAsync(
        string sourceSmtp,
        string globalAppointmentId,
        string entryId,
        string recipient,
        CancellationToken cancellationToken = default)
        => RunAsync(
            [
                "--forward-spike",
                "--source-smtp", sourceSmtp,
                "--global-id", globalAppointmentId,
                "--prepare-calendar-recipient",
                "--entry-id", entryId,
                "--to", recipient,
                "--include-details",
            ],
            cancellationToken);

    private static async Task<HostCommandResult> RunAsync(
        IReadOnlyList<string> arguments,
        CancellationToken cancellationToken)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = ResolveHostPath(),
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
        };

        foreach (var argument in arguments)
        {
            startInfo.ArgumentList.Add(argument);
        }

        using var process = new Process { StartInfo = startInfo };
        if (!process.Start())
        {
            throw new InvalidOperationException("OutlookHost could not be started.");
        }

        var standardOutputTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
        var standardErrorTask = process.StandardError.ReadToEndAsync(cancellationToken);

        try
        {
            await process.WaitForExitAsync(cancellationToken).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            TryKill(process);
            throw;
        }

        var standardOutput = await standardOutputTask.ConfigureAwait(false);
        var standardError = await standardErrorTask.ConfigureAwait(false);
        return new HostCommandResult(process.ExitCode, standardOutput.Trim(), standardError.Trim());
    }

    private static string ResolveHostPath()
    {
        var overridePath = Environment.GetEnvironmentVariable("OUTLOOK_ALIGNER_HOST_PATH");
        if (!string.IsNullOrWhiteSpace(overridePath) && File.Exists(overridePath))
        {
            return overridePath;
        }

        var siblingPath = Path.Combine(AppContext.BaseDirectory, "OutlookAligner.OutlookHost.exe");
        if (File.Exists(siblingPath))
        {
            return siblingPath;
        }

        throw new FileNotFoundException(
            "OutlookAligner.OutlookHost.exe was not found beside the UI. Set OUTLOOK_ALIGNER_HOST_PATH for development builds.",
            siblingPath);
    }

    private static void TryKill(Process process)
    {
        try
        {
            if (!process.HasExited)
            {
                process.Kill(entireProcessTree: true);
            }
        }
        catch (InvalidOperationException)
        {
        }
    }
}

internal sealed record HostCommandResult(int ExitCode, string StandardOutput, string StandardError)
{
    internal bool Success => ExitCode == 0;

    internal string DiagnosticText
        => string.IsNullOrWhiteSpace(StandardError)
            ? StandardOutput
            : string.IsNullOrWhiteSpace(StandardOutput)
                ? StandardError
                : StandardOutput + Environment.NewLine + StandardError;

    internal void ThrowIfFailed(string operation)
    {
        if (Success)
        {
            return;
        }

        var detail = string.IsNullOrWhiteSpace(DiagnosticText)
            ? $"OutlookHost exit code {ExitCode}."
            : DiagnosticText;
        throw new InvalidOperationException($"{operation} failed: {detail}");
    }
}
