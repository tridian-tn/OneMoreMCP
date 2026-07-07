namespace OneMoreMcp.Core;

/// <summary>The outcome of one <c>OneMoreCli.exe</c> invocation.</summary>
/// <param name="ExitCode">Process exit code — <c>0</c> success, <c>1</c> error (per the CLI contract).</param>
/// <param name="StdOut">Captured standard output (XML/JSON/text, depending on the command).</param>
/// <param name="StdErr">Captured standard error.</param>
public sealed record CliResult(int ExitCode, string StdOut, string StdErr)
{
    /// <summary>True when the CLI reported success (exit code 0).</summary>
    public bool Ok => ExitCode == 0;
}
