namespace OneMoreMcp.Core;

/// <summary>
/// Thrown when a OneMore CLI invocation fails — a non-zero exit code, a missing executable, a
/// timeout, or a cancellation. Carries the command line and captured stderr so the MCP layer can
/// surface an actionable message to the caller.
/// </summary>
public sealed class CliException : Exception
{
    public CliException(string message, string? commandLine = null, int? exitCode = null, string? stdErr = null)
        : base(message)
    {
        CommandLine = commandLine;
        ExitCode = exitCode;
        StdErr = stdErr;
    }

    /// <summary>The invocation that failed (command + args), for logs and diagnostics.</summary>
    public string? CommandLine { get; }

    /// <summary>The process exit code, when the process actually ran.</summary>
    public int? ExitCode { get; }

    /// <summary>Captured standard error, when available.</summary>
    public string? StdErr { get; }
}
