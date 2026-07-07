using System.Diagnostics;
using System.Text;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using OneMoreMcp.Core;

namespace OneMoreMcp.App;

/// <summary>Executes OneMore CLI commands. Abstracted so the tool layer can be tested without the real exe.</summary>
public interface IOneMoreRunner
{
    /// <summary>Resolves the <c>OneMoreCli.exe</c> path (config override, else standard install), or null if not found.</summary>
    string? TryResolveCliPath();

    /// <summary>
    /// Runs one command to completion and returns its captured output and exit code. Invocations are
    /// serialised so at most one talks to OneNote at a time. Throws <see cref="CliException"/> when the
    /// process can't be started, times out, or is cancelled — a non-zero exit is returned, not thrown.
    /// </summary>
    Task<CliResult> RunAsync(OneMoreCommand command, CancellationToken cancellationToken = default);
}

/// <summary>
/// The concrete runner: discovers <c>OneMoreCli.exe</c>, then executes each command as a child process
/// while holding a single global lock. OneNote automation is effectively single-threaded, so serialising
/// here keeps two tool calls from colliding on the one OneNote session.
/// </summary>
public sealed class OneMoreCliRunner : IOneMoreRunner
{
    // Candidate install layouts, relative to a Program Files / LocalAppData root. The first is the
    // actual current layout; the second is what the OneMore CLI docs describe (kept for older/other builds).
    private static readonly string[][] CandidateRelativePaths =
    {
        new[] { "River", "OneMoreAddIn", "OneMoreCli.exe" },
        new[] { "River Software", "OneMore", "OneMoreCli.exe" },
    };

    private readonly IOptionsMonitor<OneMoreMcpOptions> _options;
    private readonly ILogger<OneMoreCliRunner> _log;
    private readonly SemaphoreSlim _gate = new(1, 1);

    public OneMoreCliRunner(IOptionsMonitor<OneMoreMcpOptions> options, ILogger<OneMoreCliRunner> log)
    {
        _options = options;
        _log = log;
    }

    public string? TryResolveCliPath()
    {
        // A configured path wins when it exists; otherwise fall back to scanning the standard install
        // locations, so a stale/blank CliPath (or a Program Files x86 install) still resolves.
        var configured = _options.CurrentValue.CliPath;
        if (!string.IsNullOrWhiteSpace(configured) && File.Exists(configured))
            return configured;

        foreach (var root in SearchRoots())
        {
            foreach (var relative in CandidateRelativePaths)
            {
                var candidate = Path.Combine(new[] { root }.Concat(relative).ToArray());
                if (File.Exists(candidate)) return candidate;
            }
        }
        return null;
    }

    public async Task<CliResult> RunAsync(OneMoreCommand command, CancellationToken cancellationToken = default)
    {
        var exe = TryResolveCliPath()
            ?? throw new CliException(
                "OneMore CLI not found. Install the OneMore add-in for OneNote, or set 'CliPath' in the configuration " +
                "to the full path of OneMoreCli.exe.",
                command.ToString());

        var argv = command.Build();
        var timeout = TimeSpan.FromSeconds(Math.Max(1, _options.CurrentValue.CommandTimeoutSeconds));

        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            return await ExecuteAsync(exe, argv, command, timeout, cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            _gate.Release();
        }
    }

    private async Task<CliResult> ExecuteAsync(
        string exe, IReadOnlyList<string> argv, OneMoreCommand command, TimeSpan timeout, CancellationToken cancellationToken)
    {
        var psi = new ProcessStartInfo(exe)
        {
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
            WorkingDirectory = Path.GetDirectoryName(exe) ?? Environment.CurrentDirectory,
        };
        // Skip argv[0] (the command name is passed as the first argument, not the exe name).
        foreach (var arg in argv) psi.ArgumentList.Add(arg);

        using var process = new Process { StartInfo = psi };
        var stdout = new StringBuilder();
        var stderr = new StringBuilder();
        process.OutputDataReceived += (_, e) => { if (e.Data is not null) stdout.AppendLine(e.Data); };
        process.ErrorDataReceived += (_, e) => { if (e.Data is not null) stderr.AppendLine(e.Data); };

        _log.LogInformation("Running OneMore CLI: {Command}", command.ToString());

        try
        {
            process.Start();
        }
        catch (Exception ex)
        {
            throw new CliException($"Could not start OneMore CLI: {ex.Message}", command.ToString(), stdErr: ex.Message);
        }

        process.BeginOutputReadLine();
        process.BeginErrorReadLine();

        using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeoutCts.CancelAfter(timeout);
        try
        {
            await process.WaitForExitAsync(timeoutCts.Token).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            TryKill(process);
            var reason = cancellationToken.IsCancellationRequested
                ? "OneMore CLI call was cancelled."
                : $"OneMore CLI call timed out after {timeout.TotalSeconds:0}s.";
            throw new CliException(reason, command.ToString(), stdErr: stderr.ToString());
        }

        // Ensure the async output/error pumps have flushed before reading the buffers.
        process.WaitForExit();

        var result = new CliResult(process.ExitCode, stdout.ToString(), stderr.ToString());
        if (!result.Ok)
            _log.LogWarning("OneMore CLI exited {ExitCode}: {StdErr}", result.ExitCode, result.StdErr.Trim());
        return result;
    }

    private static void TryKill(Process process)
    {
        try { if (!process.HasExited) process.Kill(entireProcessTree: true); }
        catch { /* best effort — the process may have exited between the check and the kill */ }
    }

    private static IEnumerable<string> SearchRoots()
    {
        var folders = new[]
        {
            Environment.SpecialFolder.ProgramFiles,
            Environment.SpecialFolder.ProgramFilesX86,
            Environment.SpecialFolder.LocalApplicationData, // some add-in installs land here
        };
        foreach (var folder in folders)
        {
            var path = Environment.GetFolderPath(folder);
            if (!string.IsNullOrEmpty(path)) yield return path;
        }
    }
}
