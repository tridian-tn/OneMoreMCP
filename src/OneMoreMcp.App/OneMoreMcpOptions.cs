namespace OneMoreMcp.App;

/// <summary>Configuration for the MCP server, bound from the "OneMoreMcp" section.</summary>
public sealed class OneMoreMcpOptions
{
    public const string SectionName = "OneMoreMcp";

    /// <summary>Loopback TCP port the MCP endpoint listens on.</summary>
    public int Port { get; set; } = 3002;

    /// <summary>
    /// Serve over HTTPS. Off by default: the endpoint is loopback-only, so plain HTTP never leaves
    /// the machine and avoids the certificate step. When true the app uses a self-signed localhost
    /// certificate; when false it serves plain HTTP.
    /// </summary>
    public bool UseHttps { get; set; } = false;

    /// <summary>
    /// Install the self-signed certificate into the current user's Trusted Root store so Claude
    /// accepts it. The first install shows a one-time Windows consent prompt (no admin needed).
    /// </summary>
    public bool TrustCertificate { get; set; } = true;

    /// <summary>
    /// Full path to <c>OneMoreCli.exe</c>. Defaults to the conventional install location
    /// (<see cref="DefaultCliPath"/>). If the configured path doesn't exist, the runner falls back to
    /// auto-detecting the standard locations (Program Files and Program Files x86); blank does the same.
    /// </summary>
    public string CliPath { get; set; } = DefaultCliPath;

    /// <summary>The conventional OneMore CLI install location: <c>%ProgramFiles%\River\OneMoreAddIn\OneMoreCli.exe</c>.</summary>
    public static string DefaultCliPath { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
        "River", "OneMoreAddIn", "OneMoreCli.exe");

    /// <summary>
    /// Master gate for content-changing tools (create/overwrite pages, export, hashtags, cleanup).
    /// Off by default so a fresh install is read-only. The append-only <c>append_to_page</c> tool is
    /// intentionally exempt — it can only add text, never overwrite, and never exposes page content.
    /// </summary>
    public bool AllowWrites { get; set; } = false;

    /// <summary>Default rendering for read tools that return page/hierarchy content: <c>markdown</c> or <c>xml</c>.</summary>
    public string DefaultFormat { get; set; } = "markdown";

    /// <summary>
    /// How long a single CLI invocation may run before it's killed. OneMore operations spanning many
    /// pages can take a minute or two, so this is generous by default.
    /// </summary>
    public int CommandTimeoutSeconds { get; set; } = 150;

    /// <summary>
    /// Optional folder that <c>export</c> output paths must stay within. Blank allows any path; set it
    /// to confine exports (and the files the LLM can create) to one directory.
    /// </summary>
    public string ExportRoot { get; set; } = "";
}
